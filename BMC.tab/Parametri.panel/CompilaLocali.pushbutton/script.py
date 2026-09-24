"""Fill the empty UTI parameters of the rooms (PGI R06 3.11.12, PIR R06 LOC) from level, room number and the nd rule"""

__title__ = "Compila\nlocali"
__author__ = "Luca Rosati"
__guida__ = 'Compila i parametri UTI vuoti dei locali: edificio e piano dal livello, abbreviazione dal numero, nd per appartamento, reparto e scala.'

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# Only empty values are written. Groups (each one can be excluded in the preview):
#   edificio      UTI_TE_Edificio from the level name (E0C_AR_L01A_pf_+0.59 -> E0C)
#   piano         UTI_TE_Piano from the level name (-> L01A)
#   abbreviazione UTI_TE_Abbreviazione Locale from the number PIANO.FUNZIONE.NN (-> BG), else PGI table
#   appartamento  UTI_TE_Tipologia / Codice appartamento = nd (non residential model, riunione 21/09 nd rule)
#   reparto       UTI_TE_Codice reparto = nd until the department codes are defined
#   scala         UTI_TE_Scala = nd until the stair codes are defined
# Rooms whose number does not start with their level code are reported (not changed).

from collections import Counter

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script

import bmc_rules as R
from bmc_utils import commit, idv, instances, level_name, param_state, save_log

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

GROUPS = [
    ("edificio", "UTI_TE_Edificio", "dal nome del livello"),
    ("piano", "UTI_TE_Piano", "dal nome del livello"),
    ("abbreviazione", "UTI_TE_Abbreviazione Locale", "dal Numero PIANO.FUNZIONE.NN"),
    ("appartamento", "UTI_TE_Tipologia appartamento", "nd (modello non residenziale)"),
    ("appartamento", "UTI_TE_Codice appartamento", "nd (modello non residenziale)"),
    ("reparto", "UTI_TE_Codice reparto", "nd (codici reparto da definire)"),
    ("scala", "UTI_TE_Scala", "nd (codici vano scala da definire)"),
]


def room_text(room, bip):
    p = room.get_Parameter(bip)
    return (p.AsString() or "").strip() if p is not None else ""


# ---------------------------------------------------------------- plan (read only)

changes = []  # (room, group, parameter, new)
missing_params = set()
mismatch, unresolved = [], Counter()
for room in instances(doc, [BuiltInCategory.OST_Rooms]):
    if room.Location is None:
        continue  # not placed: reported by the checker 07.1
    number = room_text(room, BuiltInParameter.ROOM_NUMBER)
    name = room_text(room, BuiltInParameter.ROOM_NAME)
    m = R.RE_LEVEL.match(level_name(doc, room) or "")
    if m and number and not number.startswith(m.group(2)):
        mismatch.append((room, number, m.group(2)))
    proposed = {
        "UTI_TE_Edificio": m.group(1) if m else None,
        "UTI_TE_Piano": m.group(2) if m else None,
        "UTI_TE_Abbreviazione Locale": R.room_abbreviation(number, name),
        "UTI_TE_Tipologia appartamento": R.ND,
        "UTI_TE_Codice appartamento": R.ND,
        "UTI_TE_Codice reparto": R.ND,
        "UTI_TE_Scala": R.ND,
    }
    for group, param, _ in GROUPS:
        state, _value = param_state(room, param)
        if state == "missing":
            missing_params.add(param)
        elif state == "empty":
            if proposed[param]:
                changes.append((room, group, param, proposed[param]))
            else:
                unresolved[param] += 1

# ---------------------------------------------------------------- preview

by_group = Counter(c[1] for c in changes)
by_param = Counter(c[2] for c in changes)
output.print_md("# BMC - Compila locali (anteprima)")
output.print_md(
    "Modello: **%s**. Nessuna modifica e' stata ancora fatta. Solo campi vuoti."
    % doc.Title
)
output.print_table(
    table_data=[[g, p, rule, by_param[p]] for g, p, rule in GROUPS],
    columns=["Gruppo", "Parametro", "Regola", "Locali da compilare"],
)
examples = {}
for room, group, param, new in changes:
    examples.setdefault(param, Counter())[new] += 1
output.print_table(
    table_data=[
        [p, ", ".join("%s (%d)" % kv for kv in c.most_common(8))]
        for p, c in sorted(examples.items())
    ],
    columns=["Parametro", "Valori proposti"],
)
if missing_params:
    output.print_md(
        "Parametri non presenti sui locali (Parametri > Riemetti parametri R06): %s"
        % ", ".join(sorted(missing_params))
    )
if unresolved:
    output.print_md(
        "Non ricavabili: %s" % ", ".join("%s %d" % kv for kv in unresolved.items())
    )
if mismatch:
    output.print_md("## Numero locale diverso dal livello (non modificati)")
    for room, number, code in mismatch:
        output.print_md(
            "- `%s` su livello `%s` %s" % (number, code, output.linkify(room.Id))
        )

choices = [
    "%s (%d)" % (g, by_group[g])
    for g in ("edificio", "piano", "abbreviazione", "appartamento", "reparto", "scala")
    if by_group[g]
]
if not choices:
    forms.alert(
        "Parametri UTI dei locali gia' compilati: niente da fare.", exitscript=True
    )
chosen = forms.SelectFromList.show(
    choices, title="Gruppi da compilare", multiselect=True, button_name="Applica"
)
if not chosen:
    forms.alert(
        "Nessun gruppo selezionato: il modello non e' stato modificato.",
        exitscript=True,
    )
chosen = set(c.rsplit(" (", 1)[0] for c in chosen)
if not forms.alert(
    "Compilare i parametri UTI selezionati?\nL'operazione si annulla con Ctrl+Z.",
    yes=True,
    no=True,
):
    script.exit()

# ---------------------------------------------------------------- apply

log = {"model": doc.Title, "applied": Counter(), "failed": []}
t = Transaction(doc, "BMC - Compila locali")
t.Start()
try:
    for room, group, param, new in changes:
        if group not in chosen:
            continue
        try:
            room.LookupParameter(param).Set(new)
            log["applied"][param] += 1
        except Exception as ex:
            log["failed"].append([idv(room.Id), param, new, str(ex)])
except Exception as ex:
    t.RollBack()
    forms.alert(
        "Errore imprevisto, nessuna modifica applicata:\n%s" % ex, exitscript=True
    )
else:
    log["commit_status"] = commit(t)

log["applied"] = dict(log["applied"])
log_path = save_log("BMC_CompilaLocali", log)
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
    table_data=sorted(log["applied"].items()), columns=["Parametro", "Locali"]
)
if log["failed"]:
    output.print_table(
        table_data=log["failed"], columns=["Locale", "Parametro", "Valore", "Errore"]
    )
output.print_md("Log: `%s`. Poi: Model Checker (06.LOC, 07.x)." % log_path)
