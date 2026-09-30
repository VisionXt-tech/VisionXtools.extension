"""Fix BMC naming and derived parameters from the real model data: stratigraphic type names, Keynote, WBS concatenation, HOST codes, Contrassegno univoco, Si/No"""

__title__ = "Correggi\nNomenclature"
__author__ = "Luca Rosati"
__guida__ = 'Rinomina i tipi stratigrafici in base agli strati reali, compila Keynote, WBS_TE_WBS, codici HOST, Contrassegno univoco e Si/No vuoti, rinomina i livelli COO.'

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
from bmc_utils import ask_preview_only, end_preview, host_codes, instances, param_state_any

HAS_DATA = R.load_data(required=False)  # WBS R06 domain, only for WBS_TE_WBS

AA_PGI = R.STRAT_CODES
CAT_CODE = R.STRAT_CAT_CODE  # walls, floors (COP when CS), ceilings; roofs are remodelled as floors: not renamed
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


PREVIEW_ONLY = ask_preview_only()

# ---------------------------------------------------------------- plan (read only)

renames, keynotes, wbs = [], [], []
hosts = []  # (element, parameter, old, new)
uniques = []  # (element, old, new, level name)
yesno = []  # (element or type, parameter)
skipped = []  # (group, element, reason)


def plan_renames():
    used = set()
    for t in FilteredElementCollector(doc).WhereElementIsElementType():
        if isinstance(t, HostObjAttributes):
            used.add(
                (idv(t.Category.Id) if t.Category else None, t.FamilyName, ename(t))
            )
    targets = set()
    for bic in CAT_CODE:
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
                        "codice AA '%s' non previsto dal PGI (coperture: CS, nome AR_COP_CS..., referenti 25/09)"
                        % m.group(1),
                    )
                )
                continue
            cat_code = R.strat_cat_code(bic, m.group(1))  # AR_COP_CS... for roofs as floors
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
    # referents 25/09/2026: HOST = Type Mark of the host, HOST+thickness = TypeMark_thickness (PV2.01K_112).
    # Values not in that format (e.g. OLD / OLD_160 from OLD_ walls) are not admissible: they are replaced
    # as soon as the host has a valid Type Mark; valid values that differ are only reported.
    bics = [
        BuiltInCategory.OST_Doors,
        BuiltInCategory.OST_Windows,
        BuiltInCategory.OST_CurtainWallPanels,
    ]
    formats = {HOST_PARAM: RE_TYPE_MARK, R.P_HOST_THK: R.RE_HOST_THK}
    for bic in bics:
        for el in (
            FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType()
        ):
            if getattr(el, "Host", None) is None:
                continue
            tm, tm_thk = host_codes(doc, el)
            for name, expected in ((HOST_PARAM, tm), (R.P_HOST_THK, tm_thk)):
                p, current = text_param(el, name)
                if p is None or p.IsReadOnly:
                    continue
                if not RE_TYPE_MARK.match(tm or ""):
                    reason = "host senza Type Mark valido ('%s'): tipo OLD_ da sostituire o Type Mark da compilare" % tm
                    skipped.append(("Codice HOST", "id %s" % idv(el.Id), reason))
                elif not expected:
                    skipped.append(("Codice HOST", "id %s" % idv(el.Id), "l'host non ha strati ne' spessore nel nome"))
                elif not current or (current != expected and not formats[name].match(current)):
                    hosts.append((el, name, current, expected))
                elif current != expected:
                    skipped.append(
                        (
                            "Codice HOST",
                            "id %s" % idv(el.Id),
                            "valore '%s' diverso da '%s': non sovrascritto"
                            % (current, expected),
                        )
                    )


def element_xy(el):
    loc = el.Location
    if isinstance(loc, LocationPoint):
        return loc.Point.X, loc.Point.Y
    bb = el.get_BoundingBox(None)
    if bb is None:
        return None
    return (bb.Min.X + bb.Max.X) / 2.0, (bb.Min.Y + bb.Max.Y) / 2.0


def plan_uniques():
    # referents 25/09/2026: prefix + progressive number, P01... on doors (F01... on windows).
    # Order (decision 25/09/2026): levels bottom-up, then clockwise from north on each level (bmc_rules.clockwise_order).
    # The whole category is renumbered in that order: values already P/F that end up in a different place change,
    # values in another format are reported and left out of the numbering.
    for bic, prefix in R.UNIQUE_PREFIX.items():
        by_level = {}
        for el in FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType():
            p, current = text_param(el, R.P_UNIQUE)
            if p is None or p.IsReadOnly or el.SuperComponent is not None:
                continue  # nested shared components are part of their parent door/window
            m = R.RE_UNIQUE.match(current)
            xy = element_xy(el)
            if current and not (m and m.group(1) == prefix):
                reason = "valore '%s' non conforme: non sovrascritto" % current
            elif xy is None:
                reason = "elemento senza posizione: non numerato"
            else:
                level = doc.GetElement(el.LevelId)
                key = (level.Elevation if level is not None else 0.0, idv(el.LevelId))
                by_level.setdefault(key, []).append((xy[0], xy[1], (el, current, level)))
                continue
            skipped.append(("Contrassegno univoco", "id %s" % idv(el.Id), reason))
        n = 0
        for key in sorted(by_level):
            for el, current, level in R.clockwise_order(by_level[key]):
                n += 1
                new = R.unique_mark(prefix, n)
                if current != new:
                    uniques.append((el, current, new, ename(level) if level is not None else "-"))


def plan_yesno():
    # referents 25/09/2026: PIR Si/No parameters never set -> "No"; values already set are kept
    seen = set()
    for _code, _label, bics, _kind, params in R.PIR_GROUPS:
        names = [(n, lv) for n, lv in params if R.YESNO_TAG in n]
        if not names:
            continue
        for el in instances(doc, bics):
            for name, level in names:
                target, state, _ = param_state_any(doc, el, name, level)
                key = (idv(target.Id), name)
                if state != "empty" or key in seen:
                    continue
                p = target.LookupParameter(name)
                if p is None or p.IsReadOnly or p.StorageType != StorageType.Integer:
                    continue
                seen.add(key)
                yesno.append((target, name))


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
plan_uniques()
plan_yesno()
plan_datums()

# ---------------------------------------------------------------- preview

output.print_md("# BMC - Correzione nomenclature (anteprima)")
output.print_md("Modello: **%s**. Nessuna modifica e' stata ancora fatta." % doc.Title)
output.print_table(
    table_data=[
        ["Nomi dei tipi stratigrafici (da strati reali)", len(renames)],
        ["Keynote vuote -> AR_CAT_TypeMark", len(keynotes)],
        ["WBS_TE_WBS vuoto -> concatenazione dei livelli", len(wbs)],
        ["Codici HOST vuoti -> Type Mark (e spessore) dell'host", len(hosts)],
        ["Contrassegno univoco -> P01 / F01 per livello, da nord in senso orario", len(uniques)],
        ["Si/No non compilati -> No", len(yesno)],
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
        table_data=[[output.linkify(el.Id), name, new] for el, name, _, new in hosts[:MAX_ROWS]],
        columns=["Elemento", "Parametro", "Nuovo valore"],
    )
if uniques:
    output.print_md(
        "## Contrassegno univoco (prime %d righe su %d)"
        % (min(MAX_ROWS, len(uniques)), len(uniques))
    )
    output.print_table(
        table_data=[
            [output.linkify(el.Id), el.Category.Name, lv, old or "-", new]
            for el, old, new, lv in uniques[:MAX_ROWS]
        ],
        columns=["Elemento", "Categoria", "Livello", "Attuale", "Nuovo"],
    )
if yesno:
    output.print_md("## Si/No non compilati -> No")
    counts = {}
    for _el, name in yesno:
        counts[name] = counts.get(name, 0) + 1
    output.print_table(
        table_data=[[n, k] for n, k in sorted(counts.items())],
        columns=["Parametro", "Elementi/tipi"],
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

end_preview(output, PREVIEW_ONLY)

groups = []
if renames:
    groups.append("Nomi dei tipi stratigrafici (%d)" % len(renames))
if keynotes:
    groups.append("Keynote (%d)" % len(keynotes))
if wbs:
    groups.append("WBS_TE_WBS (%d)" % len(wbs))
if hosts:
    groups.append("Codice HOST (%d)" % len(hosts))
if uniques:
    groups.append("Contrassegno univoco (%d)" % len(uniques))
if yesno:
    groups.append("Si/No -> No (%d)" % len(yesno))
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
        for element, name, old, new in hosts:
            try:
                element.LookupParameter(name).Set(new)
                record(name, element, old, new)
            except Exception as ex:
                record(name, element, old, new, str(ex))
    if any(c.startswith("Contrassegno") for c in chosen):
        for element, old, new, _lv in uniques:
            try:
                element.LookupParameter(R.P_UNIQUE).Set(new)
                record("contrassegno_univoco", element, old, new)
            except Exception as ex:
                record("contrassegno_univoco", element, old, new, str(ex))
    if any(c.startswith("Si/No") for c in chosen):
        for element, name in yesno:
            try:
                element.LookupParameter(name).Set(0)
                record(name, element, None, 0)
            except Exception as ex:
                record(name, element, None, 0, str(ex))
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
