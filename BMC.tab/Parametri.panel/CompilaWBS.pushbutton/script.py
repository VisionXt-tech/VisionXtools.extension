"""Fill the empty or not admissible WBS levels (WBS R06) from the model itself and rebuild WBS_TE_WBS"""

__title__ = "Compila\nWBS"
__author__ = "Luca Rosati"
__guida__ = "Compila i livelli WBS vuoti o non ammessi ricavandoli dal modello (costanti, livello, tipo, categoria) e ricalcola WBS_TE_WBS. Non cambia i valori gia' validi."

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# Proposal for every element of the PIR groups (MUR..OGG levels 1-8 + concatenation, LOC levels 1-5):
#   L1 L2 L3 L5  most used admissible value in the model (BMA, FP, AUL, E0C for this model)
#   L6           from the element level (bmc_rules.wbs_l6_from_level: B01A -> i01, L01A -> L01)
#   L4           from L6 (interrato BG, fuori terra AG)
#   L7 L8        most used admissible pair among the instances of the same type, else of the same category
#                (walls: same AA code), only when that pair is in bmc_rules.WBS_EXPECTED; else the first
#                expected pair (source "regola", to check)
# Valid current values are never changed (bmc_rules.wbs_merge); elements with nd are skipped.
# Preview per type -> choose the sources -> one Transaction (Ctrl+Z) -> JSON log + CSV of the type table.

import codecs
import datetime
import os
from collections import Counter

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script

import bmc_rules as R
from bmc_utils import (
    commit,
    ename,
    idv,
    instances,
    layer_codes,
    level_name,
    param_state,
    save_log,
    type_mark,
)

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()
R.load_data()  # client shared parameter TXTs + WBS R06: asks the folder the first time

MAX_ROWS = 150
SOURCES = [
    ("costante", "L1-L3, L5 dal valore piu' usato nel modello"),
    ("livello", "L4, L6 dal livello dell'elemento"),
    ("tipo", "L7-L8 dagli altri elementi dello stesso tipo"),
    ("categoria", "L7-L8 dagli elementi della stessa categoria"),
    ("regola", "L7-L8 dalla tabella bmc_rules.WBS_EXPECTED (proposta, da validare)"),
    ("difforme", "L7-L8 esistenti validi ma diversi dalla regola: sostituiti con la regola"),
    ("concatenato", "WBS_TE_WBS ricalcolato"),
]

# ---------------------------------------------------------------- collect (read only)

rows = []  # dict(el, group, n_levels, current[8], level)
for code, _label, bics, kind, params in R.PIR_GROUPS:
    names = R.pir_names(params)
    if R.WBS_LEVELS[0] not in names:
        continue
    n_levels = len([n for n in R.WBS_LEVELS if n in names])
    for el in instances(doc, bics):
        if kind:
            want = WallKind.Curtain if kind == "curtain" else WallKind.Basic
            if not isinstance(el, Wall) or el.WallType.Kind != want:
                continue
        current = []
        for name in R.WBS_LEVELS[:n_levels]:
            state, value = param_state(el, name)
            current.append(value if state in ("value", "nd") else "")
        rows.append(
            {
                "el": el,
                "group": code,
                "n": n_levels,
                "current": current,
                "level": level_name(doc, el),
            }
        )

# most used admissible values: L1, L2, L3, L5 over the model; (L7, L8) per type and per category
constants = [Counter() for _ in range(8)]
by_type, by_category = {}, {}


def category_key(row):
    el = row["el"]
    aa = ""
    if row["group"] == "MUR":
        m = R.RE_TYPE_MARK.match(type_mark(doc.GetElement(el.GetTypeId())) or "")
        aa = m.group(1) if m else ""
    return (row["group"], aa)


for row in rows:
    cur = row["current"]
    for i in (0, 1, 2, 4):
        if i < len(cur) and cur[i] in R.WBS_DOMAIN[i]:
            constants[i][cur[i]] += 1
    if row["n"] == 8 and not R.wbs_level_errors(cur) and cur[6] and cur[7]:
        pair = (cur[6], cur[7])
        by_type.setdefault(idv(row["el"].GetTypeId()), Counter())[pair] += 1
        by_category.setdefault(category_key(row), Counter())[pair] += 1
constant = [c.most_common(1)[0][0] if c else "" for c in constants]


def proposal(row):
    """(proposed values, {level index: source}, expected (L7, L8) pairs)."""
    expected = []
    l6 = R.wbs_l6_from_level(row["level"])
    values = [
        constant[0],
        constant[1],
        constant[2],
        R.wbs_l4_from_l6(l6) or "",
        constant[4],
        l6 or "",
        "",
        "",
    ]
    source = {
        0: "costante",
        1: "costante",
        2: "costante",
        3: "livello",
        4: "costante",
        5: "livello",
    }
    if row["n"] == 8:
        tid = idv(row["el"].GetTypeId())
        key = category_key(row)
        el_type = doc.GetElement(row["el"].GetTypeId())
        expected = R.wbs_expected(
            row["group"],
            row["el"].Category.BuiltInCategory,
            type_mark(el_type) if el_type else "",
            layer_codes(doc, el_type) if el_type else [],
            ename(el_type) if el_type else "",
        )

        def usable(counter):
            # the most used pair of the type/category, only if it is one of the expected pairs:
            # a wrong pair already in the model must not be copied to other elements
            if not counter:
                return None
            best = counter.most_common(1)[0][0]
            return best if not expected or best in expected else None

        pair, src = usable(by_type.get(tid)), "tipo"
        if not pair:
            pair, src = usable(by_category.get(key)) or usable(by_category.get((key[0], ""))), "categoria"
        if not pair:
            pair, src = (expected[0] if expected else None), "regola"
        if pair:
            values[6], values[7] = pair
            source[6] = source[7] = src
    return values[: row["n"]], source, (expected if row["n"] == 8 else [])


# ---------------------------------------------------------------- plan

changes = []  # (row, level index, old, new, source)
concat = []  # (row, old, new)
skipped = Counter()
per_type = {}  # type id -> dict(name, category, n, pair, source)
for row in rows:
    cur = row["current"]
    if R.ND in cur:
        skipped["nd presente: gestione manuale"] += 1
        continue
    proposed, source, expected = proposal(row)
    final = R.wbs_merge(cur, proposed + [""] * (8 - row["n"]), row["n"])
    if expected and not R.wbs_level_errors(final) and final[6] and (final[6], final[7]) not in expected:
        # valid but not the expected pair for this kind of element (e.g. curtain wall on 402 opaque facades)
        final[6], final[7] = expected[0]
        source[6] = source[7] = "difforme"
    changed = False
    for i in range(row["n"]):
        if final[i] != cur[i]:
            changes.append((row, i, cur[i], final[i], source.get(i, "?")))
            changed = True
    if row["n"] == 8:
        remaining = R.wbs_level_errors(final) + [
            "L%d vuoto" % (i + 1) for i in range(8) if not final[i]
        ]
        if remaining:
            skipped["non risolto: " + remaining[0].split(" `")[0]] += 1
        else:
            state, old = param_state(row["el"], R.WBS_CONCAT)
            new = R.WBS_SEP.join(final)
            if state != "missing" and old != new:
                concat.append((row, old or "", new))
    if changed and row["n"] == 8:
        t = doc.GetElement(row["el"].GetTypeId())
        info = per_type.setdefault(
            idv(row["el"].GetTypeId()),
            {
                "name": ename(t) if t else "-",
                "category": row["el"].Category.Name,
                "n": 0,
                "pair": "",
                "source": "",
            },
        )
        info["n"] += 1
        info["pair"] = "%s / %s" % (final[6], final[7])
        info["source"] = source.get(6, "")

# ---------------------------------------------------------------- preview

count_by_source = Counter(c[4] for c in changes)
count_by_source["concatenato"] = len(concat)
output.print_md("# BMC - Compila WBS (anteprima)")
output.print_md(
    "Modello: **%s**. Nessuna modifica e' stata ancora fatta. Costanti dal modello: L1 `%s`, L2 `%s`, L3 `%s`, L5 `%s`."
    % (doc.Title, constant[0], constant[1], constant[2], constant[4])
)
output.print_table(
    table_data=[[key, label, count_by_source[key]] for key, label in SOURCES],
    columns=["Fonte", "Descrizione", "Valori da scrivere"],
)
if per_type:
    output.print_md("## L7 / L8 proposti per tipo")
    table = sorted(
        per_type.values(),
        key=lambda r: (r["source"] != "regola", r["category"], r["name"]),
    )
    output.print_table(
        table_data=[
            [r["category"], r["name"], r["n"], r["pair"], r["source"]]
            for r in table[:MAX_ROWS]
        ],
        columns=["Categoria", "Tipo", "Elementi", "L7 / L8", "Fonte"],
    )
if skipped:
    output.print_md("## Non compilati")
    output.print_table(
        table_data=[[k, v] for k, v in skipped.most_common()],
        columns=["Motivo", "Elementi"],
    )

csv_path = os.path.join(
    os.path.expanduser("~"),
    "Documents",
    "BMC_CompilaWBS_tipi_%s.csv" % datetime.datetime.now().strftime("%Y%m%d_%H%M"),
)
with codecs.open(csv_path, "w", encoding="utf-8-sig") as f:
    f.write("Categoria;Tipo;Elementi;L7 / L8;Fonte\n")
    for r in sorted(per_type.values(), key=lambda r: (r["category"], r["name"])):
        f.write(
            "%s;%s;%d;%s;%s\n"
            % (
                r["category"],
                r["name"].replace(";", ","),
                r["n"],
                r["pair"],
                r["source"],
            )
        )
output.print_md("Tabella tipi (CSV per la validazione): `%s`" % csv_path)

groups = ["%s (%d)" % (k, count_by_source[k]) for k, _ in SOURCES if count_by_source[k]]
if not groups:
    forms.alert("WBS gia' compilata: niente da fare.", exitscript=True)
chosen = forms.SelectFromList.show(
    groups,
    title="Fonti da applicare (vedi anteprima)",
    multiselect=True,
    button_name="Applica",
)
if not chosen:
    forms.alert(
        "Nessuna fonte selezionata: il modello non e' stato modificato.",
        exitscript=True,
    )
chosen = set(c.rsplit(" (", 1)[0] for c in chosen)
if not forms.alert(
    "Scrivere i valori WBS delle fonti selezionate?\nL'operazione si annulla con Ctrl+Z.",
    yes=True,
    no=True,
):
    script.exit()

# ---------------------------------------------------------------- apply

log = {"model": doc.Title, "applied": Counter(), "failed": []}
t = Transaction(doc, "BMC - Compila WBS")
t.Start()
try:
    written = set()
    for row, i, old, new, src in changes:
        if src not in chosen:
            continue
        try:
            row["el"].LookupParameter(R.WBS_LEVELS[i]).Set(new)
            log["applied"][src] += 1
            written.add(idv(row["el"].Id))
        except Exception as ex:
            log["failed"].append([idv(row["el"].Id), R.WBS_LEVELS[i], new, str(ex)])
    if "concatenato" in chosen:
        for row, old, _ in concat:
            # rebuilt from the values really in the model after the writes above
            values = [param_state(row["el"], n)[1] or "" for n in R.WBS_LEVELS]
            if R.wbs_level_errors(values) or not all(values):
                continue
            try:
                row["el"].LookupParameter(R.WBS_CONCAT).Set(R.WBS_SEP.join(values))
                log["applied"]["concatenato"] += 1
            except Exception as ex:
                log["failed"].append([idv(row["el"].Id), R.WBS_CONCAT, "", str(ex)])
except Exception as ex:
    t.RollBack()
    forms.alert(
        "Errore imprevisto, nessuna modifica applicata:\n%s" % ex, exitscript=True
    )
else:
    log["commit_status"] = commit(t)

log["applied"] = dict(log["applied"])
log_path = save_log("BMC_CompilaWBS", log)
output.print_md("---")
if log["commit_status"] != "Committed":
    output.print_md(
        "# ATTENZIONE: transazione %s, il modello NON e' stato modificato"
        % log["commit_status"]
    )
output.print_md(
    "# Valori scritti: %d - non riusciti: %d"
    % (sum(log["applied"].values()), len(log["failed"]))
)
output.print_table(
    table_data=sorted(log["applied"].items()), columns=["Fonte", "Valori"]
)
if log["failed"]:
    output.print_table(
        table_data=log["failed"][:MAX_ROWS],
        columns=["Elemento", "Parametro", "Valore", "Errore"],
    )
output.print_md("Log: `%s`. Poi: Model Checker (06.WBS e 06.xxx)." % log_path)
