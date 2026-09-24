"""Fix BMC naming and derived parameters from the real model data: stratigraphic type names, Keynote, WBS concatenation, HOST code"""

__title__ = "Correggi\nNomenclature"
__author__ = "Luca Rosati"
__guida__ = 'Rinomina i tipi stratigrafici in base agli strati reali, compila Keynote, WBS_TE_WBS e codice HOST vuoti, rinomina i livelli COO.'

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# Flow: compute every change (read only) -> preview in the output window -> choose groups -> confirm ->
# one Transaction (Ctrl+Z undoes it) -> JSON log with old and new values.
# Only deterministic fixes are applied; cases that need a decision are listed and skipped.

import clr
import codecs
import datetime
import json
import os

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

import bmc_rules as R

HAS_DATA = R.load_data(required=False)  # WBS R06 domain, only for WBS_TE_WBS

AA_PGI = R.STRAT_CODES
CAT_CODE = R.STRAT_CAT_CODE  # walls, floors, ceilings (roofs are remodelled as floors: not renamed)
RE_MATERIAL = R.RE_MATERIAL
RE_TYPE_MARK = R.RE_TYPE_MARK
RE_CODED_TYPE = R.RE_CODE_PREFIX
WBS_LEVELS = R.WBS_LEVELS
WBS_CONCAT = R.WBS_CONCAT
HOST_PARAM = R.P_HOST
MAX_ROWS = 200
LEVEL_SEA = R.LEVEL_SEA


def ename(element):
    # ElementType redeclares Name as set-only: read it through Element
    return Element.Name.GetValue(element)


def idv(element_id):
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def to_mm(value_ft):
    return UnitUtils.ConvertFromInternalUnits(value_ft, UnitTypeId.Millimeters)


def text_param(element, name):
    p = element.LookupParameter(name)
    if p is None or p.StorageType != StorageType.String:
        return None, None
    return p, (p.AsString() or "").strip()


# ---------------------------------------------------------------- plan (read only)

renames, keynotes, wbs, hosts = [], [], [], []
skipped = []  # (group, element, reason)


def plan_renames():
    used = set()
    for t in FilteredElementCollector(doc).WhereElementIsElementType():
        if isinstance(t, HostObjAttributes):
            used.add(
                (idv(t.Category.Id) if t.Category else None, t.FamilyName, ename(t))
            )
    targets = set()
    for bic, cat_code in CAT_CODE.items():
        for t in (
            FilteredElementCollector(doc).OfCategory(bic).WhereElementIsElementType()
        ):
            if not isinstance(t, HostObjAttributes):
                continue
            cs = t.GetCompoundStructure()
            if cs is None:
                continue
            name = ename(t)
            if name.startswith("OLD_"):
                skipped.append(
                    ("Nome tipo", name, "tipo OLD_: va sostituito, non rinominato")
                )
                continue
            if name.startswith("WIP_"):
                # work in progress of a modeller: renaming it may duplicate an existing Type Mark
                skipped.append(("Nome tipo", name, "tipo WIP_ in lavorazione: non rinominato"))
                continue
            tm_param = t.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK)
            type_mark = tm_param.AsString() if tm_param is not None else None
            m = RE_TYPE_MARK.match(type_mark or "")
            if not m:
                skipped.append(
                    (
                        "Nome tipo",
                        name,
                        "Type Mark assente o non valido: %s" % type_mark,
                    )
                )
                continue
            if m.group(1) not in AA_PGI:
                skipped.append(
                    (
                        "Nome tipo",
                        name,
                        "codice AA '%s' non previsto dal PGI (decisione dei referenti)"
                        % m.group(1),
                    )
                )
                continue
            codes, total = [], 0.0
            ok = True
            for layer in cs.GetLayers():
                material = doc.GetElement(layer.MaterialId)
                mm = (
                    RE_MATERIAL.match(ename(material)) if material is not None else None
                )
                if mm is None:
                    ok = False
                    break
                codes.append(mm.group(1))
                total += to_mm(layer.Width)
            if not ok:
                skipped.append(
                    ("Nome tipo", name, "uno strato senza materiale codificato")
                )
                continue
            new = "AR_%s_%s_%s_%03d" % (
                cat_code,
                type_mark,
                "-".join(codes),
                # layer widths are feet: 203.5 mm sums to 203.4999..., round to 0.001 mm first (half up, like Excel)
                int(round(round(total, 3))),
            )
            if new == name:
                continue
            key = (idv(t.Category.Id), t.FamilyName, new)
            if key in used or key in targets:
                skipped.append(("Nome tipo", name, "il nome %s esiste gia'" % new))
                continue
            targets.add(key)
            renames.append((t, name, new))


def plan_keynotes():
    # Keynote follows the name the type will have after the rename
    new_names = dict((idv(t.Id), new) for t, _, new in renames)
    for t in FilteredElementCollector(doc).WhereElementIsElementType():
        if t.Category is None or t.Category.CategoryType != CategoryType.Model:
            continue
        m = RE_CODED_TYPE.match(new_names.get(idv(t.Id), ename(t)))
        if not m:
            continue
        p = t.get_Parameter(BuiltInParameter.KEYNOTE_PARAM)
        if p is None or p.IsReadOnly:
            continue
        current = (p.AsString() or "").strip()
        if not current:
            keynotes.append((t, current, m.group(1)))
        elif current != m.group(1):
            skipped.append(
                (
                    "Keynote",
                    ename(t),
                    "Keynote '%s' diversa da '%s': non sovrascritta"
                    % (current, m.group(1)),
                )
            )


def plan_wbs():
    if not HAS_DATA:
        skipped.append(("WBS_TE_WBS", "-", "cartella dati BMC non selezionata: valori WBS non verificabili, non compilato"))
        return
    for el in FilteredElementCollector(doc).WhereElementIsNotElementType():
        p_concat, current = text_param(el, WBS_CONCAT)
        if p_concat is None or p_concat.IsReadOnly:
            continue
        values = []
        for name in WBS_LEVELS:
            _, v = text_param(el, name)
            values.append(v or "")
        filled = [v for v in values if v]
        if not filled:
            continue
        last = max(i for i, v in enumerate(values) if v)
        if any(not v for v in values[: last + 1]):
            skipped.append(
                (
                    "WBS_TE_WBS",
                    "id %s" % idv(el.Id),
                    "livelli WBS con buchi: %s" % ".".join(values),
                )
            )
            continue
        # WBS R06: the concatenation is written only when every level is an admissible value
        errors = R.wbs_level_errors(values)
        if errors:
            skipped.append(("WBS_TE_WBS", "id %s" % idv(el.Id), "valori WBS non ammessi: " + "; ".join(errors)))
            continue
        new = R.WBS_SEP.join(values[: last + 1])
        # meeting 21/09/2026: if one WBS field is nd, all of them are nd
        if R.ND in values[: last + 1]:
            if any(v != R.ND for v in values[: last + 1]):
                skipped.append(
                    (
                        "WBS_TE_WBS",
                        "id %s" % idv(el.Id),
                        "livelli WBS in parte nd: %s" % new,
                    )
                )
                continue
            new = R.ND
        if not current:
            wbs.append((el, current, new))
        elif current != new:
            skipped.append(
                (
                    "WBS_TE_WBS",
                    "id %s" % idv(el.Id),
                    "valore '%s' diverso da '%s': non sovrascritto" % (current, new),
                )
            )


def plan_hosts():
    bics = [
        BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_Windows,
        BuiltInCategory.OST_CurtainWallPanels,
    ]
    for bic in bics:
        for el in (
            FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        ):
            p, current = text_param(el, HOST_PARAM)
            host = getattr(el, "Host", None)
            if p is None or p.IsReadOnly or host is None:
                continue
            host_type = doc.GetElement(host.GetTypeId())
            tm = (
                host_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK)
                if host_type is not None
                else None
            )
            expected = (tm.AsString() or "").strip() if tm is not None else ""
            if not expected:
                skipped.append(
                    ("Codice HOST", "id %s" % idv(el.Id), "l'host non ha Type Mark")
                )
                continue
            if not current:
                hosts.append((el, current, expected))
            elif current != expected:
                skipped.append(
                    (
                        "Codice HOST",
                        "id %s" % idv(el.Id),
                        "valore '%s' diverso da '%s': non sovrascritto"
                        % (current, expected),
                    )
                )


def plan_datums():
    lv_names = [lv.Name for lv in FilteredElementCollector(doc).OfClass(Level)]
    for lv in FilteredElementCollector(doc).OfClass(Level):
        if lv.Name == LEVEL_SEA or not lv.Name.startswith("COO - "):
            continue
        new = "COO_" + lv.Name[len("COO - "):].strip()
        if new in lv_names:
            skipped.append(("Livello", lv.Name, "il livello %s esiste gia'" % new))
            continue
        datums.append(("Livello", lv, lv.Name, new))


datums = []
plan_renames()
plan_keynotes()
plan_wbs()
plan_hosts()
plan_datums()

# ---------------------------------------------------------------- preview

output.print_md("# BMC - Correzione nomenclature (anteprima)")
output.print_md("Modello: **%s**. Nessuna modifica e' stata ancora fatta." % doc.Title)
output.print_table(
    table_data=[
        ["Nomi dei tipi stratigrafici (da strati reali)", len(renames)],
        ["Keynote vuote -> AR_CAT_TypeMark", len(keynotes)],
        ["WBS_TE_WBS vuoto -> concatenazione dei livelli", len(wbs)],
        ["Codice elemento tecnico HOST vuoto -> Type Mark dell'host", len(hosts)],
        ["Nomi livelli COO", len(datums)],
        ["Casi esclusi (serve una decisione)", len(skipped)],
    ],
    columns=["Correzione", "Elementi"],
)
if renames:
    output.print_md("## Nomi dei tipi")
    output.print_table(
        table_data=[[output.linkify(t.Id), old, new] for t, old, new in renames],
        columns=["Tipo", "Nome attuale", "Nuovo nome"],
    )
if keynotes:
    output.print_md("## Keynote")
    output.print_table(
        table_data=[[output.linkify(t.Id), ename(t), new] for t, _, new in keynotes],
        columns=["Tipo", "Nome tipo", "Nuova Keynote"],
    )
if wbs:
    output.print_md(
        "## WBS_TE_WBS (prime %d righe su %d)" % (min(MAX_ROWS, len(wbs)), len(wbs))
    )
    output.print_table(
        table_data=[[output.linkify(el.Id), new] for el, _, new in wbs[:MAX_ROWS]],
        columns=["Elemento", "Nuovo WBS_TE_WBS"],
    )
if hosts:
    output.print_md(
        "## Codice HOST (prime %d righe su %d)"
        % (min(MAX_ROWS, len(hosts)), len(hosts))
    )
    output.print_table(
        table_data=[[output.linkify(el.Id), new] for el, _, new in hosts[:MAX_ROWS]],
        columns=["Elemento", "Nuovo codice HOST"],
    )
if datums:
    output.print_md("## Livelli")
    output.print_table(
        table_data=[[k, old, new] for k, _, old, new in datums],
        columns=["Oggetto", "Nome attuale", "Nuovo nome"],
    )
if skipped:
    output.print_md("## Casi esclusi (restano da decidere)")
    grouped = {}
    for group, obj, reason in skipped:
        grouped.setdefault(
            (group, reason.split(":")[0] if ("diverso" in reason or "non ammessi" in reason) else reason), []
        ).append(obj)
    output.print_table(
        table_data=[
            [g, r, len(objs), ", ".join(objs[:5]) + (" ..." if len(objs) > 5 else "")]
            for (g, r), objs in sorted(grouped.items())
        ],
        columns=["Gruppo", "Motivo", "N.", "Esempi"],
    )

groups = []
if renames:
    groups.append("Nomi dei tipi stratigrafici (%d)" % len(renames))
if keynotes:
    groups.append("Keynote (%d)" % len(keynotes))
if wbs:
    groups.append("WBS_TE_WBS (%d)" % len(wbs))
if hosts:
    groups.append("Codice HOST (%d)" % len(hosts))
for _kind in ("Livello",):
    _n = len([d for d in datums if d[0] == _kind])
    if _n:
        groups.append("%s: rinomina (%d)" % (_kind, _n))
if not groups:
    forms.alert(
        "Nessuna correzione applicabile: il modello e' gia' allineato (vedi i casi esclusi nel report).",
        exitscript=True,
    )

chosen = forms.SelectFromList.show(
    groups,
    title="Correzioni da applicare (vedi anteprima)",
    multiselect=True,
    button_name="Applica le correzioni selezionate",
)
if not chosen:
    forms.alert(
        "Nessuna correzione selezionata: il modello non e' stato modificato.",
        exitscript=True,
    )
if not forms.alert(
    "Applicare %d gruppi di correzioni al modello?\nL'operazione si annulla con Ctrl+Z."
    % len(chosen),
    yes=True,
    no=True,
):
    script.exit()

# ---------------------------------------------------------------- apply

log = {
    "model": doc.Title,
    "date": datetime.datetime.now().isoformat(),
    "applied": [],
    "failed": [],
}


def record(kind, element, old, new, error=None):
    entry = {"kind": kind, "id": idv(element.Id), "old": old, "new": new}
    if error:
        entry["error"] = error
        log["failed"].append(entry)
    else:
        log["applied"].append(entry)


t = Transaction(doc, "BMC - Correzione nomenclature")
t.Start()
try:
    if any(c.startswith("Nomi") for c in chosen):
        for element, old, new in renames:
            try:
                element.Name = new
                record("nome_tipo", element, old, new)
            except Exception as ex:
                record("nome_tipo", element, old, new, str(ex))
    if any(c.startswith("Keynote") for c in chosen):
        for element, old, new in keynotes:
            try:
                element.get_Parameter(BuiltInParameter.KEYNOTE_PARAM).Set(new)
                record("keynote", element, old, new)
            except Exception as ex:
                record("keynote", element, old, new, str(ex))
    if any(c.startswith("WBS") for c in chosen):
        for element, old, new in wbs:
            try:
                element.LookupParameter(WBS_CONCAT).Set(new)
                record("wbs_te_wbs", element, old, new)
            except Exception as ex:
                record("wbs_te_wbs", element, old, new, str(ex))
    if any(c.startswith("Codice HOST") for c in chosen):
        for element, old, new in hosts:
            try:
                element.LookupParameter(HOST_PARAM).Set(new)
                record("codice_host", element, old, new)
            except Exception as ex:
                record("codice_host", element, old, new, str(ex))
    for kind, obj, old, new in datums:
        if "%s: rinomina" % kind not in " | ".join(chosen):
            continue
        try:
            obj.Name = new
            record(kind.lower(), obj, old, new)
        except Exception as ex:
            record(kind.lower(), obj, old, new, str(ex))
except Exception as ex:
    t.RollBack()
    forms.alert(
        "Errore imprevisto, nessuna modifica applicata:\n%s" % ex, exitscript=True
    )
else:
    # Commit returns RolledBack when Revit reports an error and the user cancels its dialog
    log["commit_status"] = str(t.Commit())

log_path = os.path.join(
    os.path.expanduser("~"),
    "Documents",
    "BMC_CorrezioneNomenclature_%s.json"
    % datetime.datetime.now().strftime("%Y%m%d_%H%M"),
)
# default=str: element ids are .NET Int64, not JSON-serializable in IronPython
text = json.dumps(log, indent=1, ensure_ascii=False, default=str)
with codecs.open(log_path, "w", encoding="utf-8") as f:
    f.write(text)

output.print_md("---")
if log["commit_status"] != "Committed":
    output.print_md(
        "# ATTENZIONE: transazione %s, il modello NON e' stato modificato"
        % log["commit_status"]
    )
output.print_md(
    "# Correzioni applicate: %d - non riuscite: %d"
    % (len(log["applied"]), len(log["failed"]))
)
if log["failed"]:
    output.print_table(
        table_data=[
            [output.linkify(ElementId(f["id"])), f["kind"], f["new"], f["error"]]
            for f in log["failed"][:MAX_ROWS]
        ],
        columns=["Elemento", "Tipo", "Valore", "Errore"],
    )
output.print_md("Log con valori precedenti e nuovi: `%s`" % log_path)
