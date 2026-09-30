"""Fill the fire compartment parameters of the stair / smoke filter walls and show them in the LR_VVF plans"""

__title__ = "Compartimenti\nREI scale"
__author__ = "Luca Rosati"
__guida__ = "Sui muri perimetrali dei vani scala e filtri fumo con tipologia verificata EI 60 (PV2.03E, PV2.01B, PV2.01C, PV2.04D) imposta VVF_SN_Compartimento = Si e VVF_TE_Resistenza al fuoco = EI 60; crea o aggiorna il filtro 'VVF_Muri di compartimentazione' (campitura rossa in pianta) e lo applica alle piante con LR_VVF nel nome (o al loro modello di vista se ne controlla i filtri). Prima anteprima, poi applica; log JSON."

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# Decision 29/09/2026: required class EI 60 for stairs and smoke filters; typologies verified in
# docs/20260929/Muri_REI_valutazione.md (Siniat systems + c.a. C30/37 of the IFC ST model behind the linings).
# External walls CV1.* excluded: towards the outside they do not separate compartments (to confirm with the fire
# designer). PRE_TE_Resistenza al fuoco of the six types was written via MCP on 29/09/2026.
# Walls: host walls bounding the VS / FF rooms (Room.GetBoundarySegments), same selection as Check > Muri REI.
# Filter: walls with VVF_SN_Compartimento = Si, so every future compartment wall shows up too.

import clr

clr.AddReference("RevitAPI")
from System.Collections.Generic import List
from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script

import bmc_rules as R
from bmc_utils import (
    ask_preview_only,
    commit,
    ename,
    end_preview,
    idv,
    instances,
    save_log,
    type_mark,
)

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

CLASS = "EI 60"
MARKS = ["PV2.03E", "PV2.01B", "PV2.01C", "PV2.04D"]
P_COMP, P_CLASS = "VVF_SN_Compartimento", "VVF_TE_Resistenza al fuoco"
FILTER_NAME = "VVF_Muri di compartimentazione"
VIEW_TAG = "LR_VVF"
RED = Color(255, 0, 0)

PREVIEW_ONLY = ask_preview_only()

# ---------------------------------------------------------------- compute (read only)

rooms = []
for room in instances(doc, [BuiltInCategory.OST_Rooms]):
    name = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()
    if room.Area > 0 and any(R.vvf_room_kinds(room.Number, name)):
        rooms.append(room)

walls, others = {}, {}  # wall id -> [wall, room numbers]; other type -> room numbers
options = SpatialElementBoundaryOptions()
for room in rooms:
    for loop in room.GetBoundarySegments(options) or []:
        for segment in loop:
            if segment.LinkElementId != ElementId.InvalidElementId:
                continue
            wall = doc.GetElement(segment.ElementId)
            if not isinstance(wall, Wall):
                continue
            mark = type_mark(wall.WallType)
            if mark in MARKS:
                walls.setdefault(idv(wall.Id), [wall, set()])[1].add(room.Number)
            else:
                others.setdefault(mark or ename(wall.WallType), set()).add(room.Number)

changes, missing = [], []  # (wall, parameter, old, new)
for wall, _numbers in walls.values():
    for pname, new in ((P_COMP, 1), (P_CLASS, CLASS)):
        p = wall.LookupParameter(pname)
        if p is None or p.IsReadOnly:
            missing.append((idv(wall.Id), pname))
            continue
        old = p.AsInteger() if p.StorageType == StorageType.Integer else p.AsString()
        if old != new:
            changes.append((wall, pname, old, new))

views = [
    v
    for v in FilteredElementCollector(doc).OfClass(ViewPlan)
    if not v.IsTemplate and VIEW_TAG in v.Name
]
FILTERS_PARAM = ElementId(BuiltInParameter.VIS_GRAPHICS_FILTERS)


def filter_target(view):
    """The view, or its template when the template controls V/G filters (then AddFilter on the view fails)."""
    if view.ViewTemplateId == ElementId.InvalidElementId:
        return view
    template = doc.GetElement(view.ViewTemplateId)
    if FILTERS_PARAM in list(template.GetNonControlledTemplateParameterIds()):
        return view
    return template


targets = {}
for v in views:
    t = filter_target(v)
    targets.setdefault(idv(t.Id), [t, []])[1].append(v.Name)
existing = [
    f
    for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement)
    if f.Name == FILTER_NAME
]
shared = [
    e
    for e in FilteredElementCollector(doc).OfClass(SharedParameterElement)
    if e.Name == P_COMP
]
if not shared:
    forms.alert(
        "Parametro condiviso '%s' non trovato nel modello: filtro non creabile."
        % P_COMP,
        exitscript=True,
    )

# ---------------------------------------------------------------- preview

output.print_md("# BMC - Compartimenti REI scale (%s)" % CLASS)
output.print_md(
    "%d locali VS/FF, **%d muri** con tipologia verificata, **%d valori da scrivere** (%s, %s)."
    % (len(rooms), len(walls), len(changes), P_COMP, P_CLASS)
)
by_mark = {}
for wall, numbers in walls.values():
    by_mark.setdefault(type_mark(wall.WallType), [0, set()])
    by_mark[type_mark(wall.WallType)][0] += 1
    by_mark[type_mark(wall.WallType)][1] |= numbers
output.print_table(
    table_data=[[m, n, len(r)] for m, (n, r) in sorted(by_mark.items())],
    columns=["Contrassegno tipo", "Muri", "Locali"],
)
if others:
    output.print_md("## Muri sul perimetro esclusi (tipologia non inclusa)")
    output.print_table(
        table_data=[[m, ", ".join(sorted(r))] for m, r in sorted(others.items())],
        columns=["Tipo", "Locali"],
    )
if missing:
    output.print_md(
        "**Parametro assente o non scrivibile su %d muri** (saltati): %s"
        % (len(missing), ", ".join("%s %s" % m for m in missing[:20]))
    )
output.print_md(
    "## Filtro `%s` (%s)" % (FILTER_NAME, "aggiornato" if existing else "nuovo")
)
output.print_md("Muri con %s = Si; in pianta campitura piena e linee rosse." % P_COMP)
if not views:
    output.print_md("**Nessuna pianta con '%s' nel nome.**" % VIEW_TAG)
output.print_table(
    table_data=[
        [
            ename(t),
            "vista" if not t.IsTemplate else "modello di vista (controlla i filtri)",
            ", ".join(names),
        ]
        for t, names in targets.values()
    ],
    columns=["Filtro applicato a", "Tipo", "Piante LR_VVF"],
)
end_preview(output, PREVIEW_ONLY)
if not forms.alert(
    "Scrivere %d valori su %d muri e applicare il filtro a %d viste/modelli?"
    % (len(changes), len(walls), len(targets)),
    yes=True,
    no=True,
):
    script.exit()

# ---------------------------------------------------------------- apply

log = {
    "model": doc.Title,
    "class": CLASS,
    "changes": [],
    "errors": [],
    "filter": None,
    "views": [],
}
t = Transaction(doc, "BMC - Compartimenti REI scale")
t.Start()
try:
    for wall, pname, old, new in changes:
        try:
            wall.LookupParameter(pname).Set(new)
            log["changes"].append(
                {"id": idv(wall.Id), "param": pname, "old": old, "new": new}
            )
        except Exception as ex:
            log["errors"].append({"id": idv(wall.Id), "param": pname, "error": str(ex)})

    rule = ParameterFilterRuleFactory.CreateEqualsRule(shared[0].Id, 1)
    element_filter = ElementParameterFilter(rule)
    categories = List[ElementId](
        [Category.GetCategory(doc, BuiltInCategory.OST_Walls).Id]
    )
    if existing:
        vvf_filter = existing[0]
        vvf_filter.SetCategories(categories)
        vvf_filter.SetElementFilter(element_filter)
    else:
        vvf_filter = ParameterFilterElement.Create(
            doc, FILTER_NAME, categories, element_filter
        )
    log["filter"] = {"name": FILTER_NAME, "id": idv(vvf_filter.Id), "new": not existing}

    solid = [
        fp
        for fp in FilteredElementCollector(doc).OfClass(FillPatternElement)
        if fp.GetFillPattern().IsSolidFill
    ][0]
    ogs = OverrideGraphicSettings()
    ogs.SetCutForegroundPatternId(solid.Id)
    ogs.SetCutForegroundPatternColor(RED)
    ogs.SetCutForegroundPatternVisible(True)
    ogs.SetCutLineColor(RED)
    ogs.SetProjectionLineColor(RED)
    for target, names in targets.values():
        if not target.IsFilterApplied(vvf_filter.Id):
            target.AddFilter(vvf_filter.Id)
        target.SetFilterOverrides(vvf_filter.Id, ogs)
        target.SetFilterVisibility(vvf_filter.Id, True)
        log["views"].append({"target": ename(target), "views": names})
except Exception as ex:
    t.RollBack()
    forms.alert(
        "Errore imprevisto, nessuna modifica applicata:\n%s" % ex, exitscript=True
    )
log["commit_status"] = commit(t)
log_path = save_log("BMC_CompartimentiREI", log)

output.print_md("---")
if log["commit_status"] != "Committed":
    output.print_md(
        "# ATTENZIONE: transazione %s, il modello NON e' stato modificato"
        % log["commit_status"]
    )
else:
    output.print_md(
        "# Valori scritti: %d - errori: %d - filtro applicato a %d viste/modelli"
        % (len(log["changes"]), len(log["errors"]), len(log["views"]))
    )
for e in log["errors"]:
    output.print_md("- muro %s `%s`: %s" % (e["id"], e["param"], e["error"]))
output.print_md("Log: `%s`" % log_path)
