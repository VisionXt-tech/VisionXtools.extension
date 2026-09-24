"""Export read-only model data (PIR R06 parameters, phases, worksets, levels, types, rooms, warnings, views) to JSON for the BMC model checker"""

__title__ = "Model\nData Export"
__author__ = "Luca Rosati"
__guida__ = 'Esporta in JSON i dati del modello (parametri e loro uso, tipi e stratigrafie, WBS, locali, viste, abachi, filtri) per le analisi e i report. Non modifica nulla.'

# Read-only: no Transaction is opened, the model is never modified.
# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.

import clr
import System
import codecs
import json
import os
import time
import datetime

clr.AddReference("System.Core")  # HashSet lives in System.Core on .NET Framework 4.8
from System.Collections.Generic import HashSet

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application

output = script.get_output()

import bmc_rules as R

ND_VALUES = [R.ND] + R.ND_VARIANTS
TOP_VALUES = 25
MAX_IDS = 5000

# (group code, categories, [(parameter, PIR level)]) from the PIR R06 rules in BMC.tab/lib/bmc_rules.py.
# EXTRA_ groups are not in the PIR: exported to size the remodelling (roofs, generic models, mullions).
PIR_GROUPS = [
    (code, bics, params) for code, _label, bics, _kind, params in R.PIR_GROUPS
] + [
    ("EXTRA_ROOFS", [BuiltInCategory.OST_Roofs], R.COMMON + R.PRE),
    ("EXTRA_GENERIC", [BuiltInCategory.OST_GenericModel], R.COMMON),
    ("EXTRA_MULLIONS", [BuiltInCategory.OST_CurtainWallMullions], R.COMMON),
]

# ---------------------------------------------------------------- helpers

errors = []
timings = []


def run(section, func):
    """Run one export section; a failure is logged and does not stop the export."""
    start = time.time()
    try:
        result = func()
    except Exception as ex:
        errors.append({"section": section, "error": str(ex)})
        result = None
    timings.append([section, round(time.time() - start, 2)])
    return result


def idv(element_id):
    try:
        return element_id.Value
    except AttributeError:
        return element_id.IntegerValue


def to_m(value_ft):
    return round(UnitUtils.ConvertFromInternalUnits(value_ft, UnitTypeId.Meters), 4)


def to_mm(value_ft):
    return round(
        UnitUtils.ConvertFromInternalUnits(value_ft, UnitTypeId.Millimeters), 1
    )


def ename(element):
    # ElementType redeclares Name as set-only: IronPython cannot read type.Name directly
    return Element.Name.GetValue(element)


def el_name(element_id):
    if element_id is None or element_id == ElementId.InvalidElementId:
        return None
    element = doc.GetElement(element_id)
    return ename(element) if element is not None else None


def workset_name(element):
    if not doc.IsWorkshared:
        return None
    try:
        return doc.GetWorksetTable().GetWorkset(element.WorksetId).Name
    except Exception:
        return None


def bip_str(element, bip):
    p = element.get_Parameter(bip)
    if p is None:
        return None
    if p.StorageType == StorageType.String:
        return p.AsString()
    return p.AsValueString()


def read_param(element, name):
    """Return (state, value). state: missing / unset / empty / nd / value."""
    p = element.LookupParameter(name)
    if p is None:
        return "missing", None
    if not p.HasValue:
        return "unset", None
    st = p.StorageType
    if st == StorageType.String:
        raw = p.AsString()
        if raw is None or raw.strip() == "":
            return "empty", raw
        if raw.strip() in ND_VALUES:
            return "nd", raw
        return "value", raw
    if st == StorageType.Integer:
        return "value", p.AsInteger()
    if st == StorageType.Double:
        return "value", p.AsValueString()
    if st == StorageType.ElementId:
        return "value", p.AsValueString()
    return "value", None


def read_param_any(element, name):
    """Instance first, then type. Returns (found_on, state, value)."""
    state, value = read_param(element, name)
    if state != "missing":
        return "I", state, value
    type_id = element.GetTypeId()
    if type_id is not None and type_id != ElementId.InvalidElementId:
        element_type = doc.GetElement(type_id)
        if element_type is not None:
            state, value = read_param(element_type, name)
            if state != "missing":
                return "T", state, value
    return "missing", "missing", None


def counter_add(counter, key, amount=1):
    counter[key] = counter.get(key, 0) + amount


def top(counter, limit=TOP_VALUES):
    return sorted(counter.items(), key=lambda kv: -kv[1])[:limit]


def collect_instances(bics):
    elements = []
    for bic in bics:
        elements.extend(
            FilteredElementCollector(doc)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    return elements


# ---------------------------------------------------------------- sections


def export_model_info():
    return {
        "title": doc.Title,
        "path": doc.PathName,
        "is_workshared": doc.IsWorkshared,
        "revit_version": app.VersionNumber,
        "revit_build": app.VersionBuild,
        "revit_language": str(app.Language),
        "shared_parameters_file": app.SharedParametersFilename,
        "user": app.Username,
        "exported_at": datetime.datetime.now().isoformat(),
        "file_size_mb": (
            round(os.path.getsize(doc.PathName) / 1048576.0, 1)
            if doc.PathName and os.path.exists(doc.PathName)
            else None
        ),
        "rules": "PGI R06 22/09/2026 (BMC.tab/lib/bmc_rules.py)",
    }


def export_project_parameters():
    shared_by_name = {}
    for spe in FilteredElementCollector(doc).OfClass(SharedParameterElement):
        shared_by_name.setdefault(spe.Name, []).append(str(spe.GuidValue))
    result = []
    it = doc.ParameterBindings.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        definition = it.Key
        binding = it.Current
        data_type = definition.GetDataType()
        group = definition.GetGroupTypeId()
        try:
            data_label = LabelUtils.GetLabelForSpec(data_type)
        except Exception:
            data_label = None
        try:
            group_label = LabelUtils.GetLabelForGroup(group)
        except Exception:
            group_label = None
        guid = None
        try:
            spe = doc.GetElement(definition.Id)
            if isinstance(spe, SharedParameterElement):
                guid = str(spe.GuidValue)
        except Exception:
            pass
        result.append(
            {
                "name": definition.Name,
                "binding": (
                    "Instance" if isinstance(binding, InstanceBinding) else "Type"
                ),
                "data_type": data_type.TypeId if data_type is not None else None,
                "data_type_label": data_label,
                "group": group.TypeId if group is not None else None,
                "group_label": group_label,
                "shared_guid": guid,
                "same_name_shared_guids": shared_by_name.get(definition.Name, []),
                "categories": sorted([c.Name for c in binding.Categories]),
            }
        )
    return sorted(result, key=lambda r: r["name"])


def export_parameter_usage():
    """How much each project parameter is compiled: counts per state, per category, top values.
    Used to decide which legacy (non PIR) parameters are reusable before deleting them.
    """
    params = []
    by_cat = {}  # category id -> [(index, definition, is_instance)]
    it = doc.ParameterBindings.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        is_instance = isinstance(it.Current, InstanceBinding)
        params.append(
            {
                "name": it.Key.Name,
                "binding": "Instance" if is_instance else "Type",
                "states": {},
                "valued_by_category": {},
                "values": {},
            }
        )
        for c in it.Current.Categories:
            by_cat.setdefault(idv(c.Id), []).append(
                (len(params) - 1, it.Key, is_instance)
            )
    for el in FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
        _usage_add(el, True, params, by_cat)
    for el in FilteredElementCollector(doc).WhereElementIsElementType().ToElements():
        _usage_add(el, False, params, by_cat)
    for p in params:
        p["values"] = top(p["values"])
    return params


def _usage_add(el, is_instance_pass, params, by_cat):
    if el.Category is None:
        return
    for index, definition, is_instance in by_cat.get(idv(el.Category.Id), []):
        if is_instance != is_instance_pass:
            continue
        p = el.get_Parameter(definition)
        if p is None:
            continue
        entry = params[index]
        if not p.HasValue:
            state, value = "unset", None
        elif p.StorageType == StorageType.String:
            value = (p.AsString() or "").strip()
            state = "value" if value else "empty"
            if value in ND_VALUES:
                state = "nd"
        elif p.StorageType == StorageType.Integer:
            state, value = "value", p.AsInteger()
        else:
            state, value = "value", p.AsValueString()
        counter_add(entry["states"], state)
        if state in ("value", "nd"):
            counter_add(entry["valued_by_category"], el.Category.Name)
            counter_add(entry["values"], value)


def export_phases():
    phases = [{"name": p.Name, "id": idv(p.Id)} for p in doc.Phases]
    statuses = [
        ("New", ElementOnPhaseStatus.New),
        ("Existing", ElementOnPhaseStatus.Existing),
        ("Demolished", ElementOnPhaseStatus.Demolished),
        ("Temporary", ElementOnPhaseStatus.Temporary),
    ]
    filters = []
    for pf in FilteredElementCollector(doc).OfClass(PhaseFilter):
        row = {"name": pf.Name}
        for label, status in statuses:
            row[label] = str(pf.GetPhaseStatusPresentation(status))
        filters.append(row)
    return {"phases": phases, "phase_filters": sorted(filters, key=lambda r: r["name"])}


def export_worksets():
    if not doc.IsWorkshared:
        return []
    return sorted(
        [
            {"name": w.Name, "id": w.Id.IntegerValue}
            for w in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
        ],
        key=lambda r: r["name"],
    )


def export_datums():
    levels = []
    for lv in FilteredElementCollector(doc).OfClass(Level):
        structural = lv.get_Parameter(BuiltInParameter.LEVEL_IS_STRUCTURAL)
        scope = lv.get_Parameter(BuiltInParameter.DATUM_VOLUME_OF_INTEREST)
        levels.append(
            {
                "name": lv.Name,
                "id": idv(lv.Id),
                "elevation_m": to_m(lv.Elevation),
                "project_elevation_m": to_m(lv.ProjectElevation),
                "structural": (
                    structural.AsInteger() if structural is not None else None
                ),
                "monitored": lv.IsMonitoringLinkElement(),
                "pinned": lv.Pinned,
                "workset": workset_name(lv),
                "scope_box": (
                    el_name(scope.AsElementId()) if scope is not None else None
                ),
            }
        )
    grids = []
    for gr in FilteredElementCollector(doc).OfClass(Grid):
        scope = gr.get_Parameter(BuiltInParameter.DATUM_VOLUME_OF_INTEREST)
        grids.append(
            {
                "name": gr.Name,
                "id": idv(gr.Id),
                "monitored": gr.IsMonitoringLinkElement(),
                "pinned": gr.Pinned,
                "workset": workset_name(gr),
                "scope_box": (
                    el_name(scope.AsElementId()) if scope is not None else None
                ),
            }
        )
    return {
        "levels": sorted(levels, key=lambda r: r["elevation_m"]),
        "grids": sorted(grids, key=lambda r: r["name"]),
    }


def export_base_points():
    result = {}
    for label, bp in [
        ("project_base_point", BasePoint.GetProjectBasePoint(doc)),
        ("survey_point", BasePoint.GetSurveyPoint(doc)),
    ]:
        row = {
            "position_m": [
                to_m(bp.Position.X),
                to_m(bp.Position.Y),
                to_m(bp.Position.Z),
            ],
            "shared_position_m": [
                to_m(bp.SharedPosition.X),
                to_m(bp.SharedPosition.Y),
                to_m(bp.SharedPosition.Z),
            ],
            "pinned": bp.Pinned,
            "clipped": bp.Clipped,
        }
        for key, bip in [
            ("north_south", BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM),
            ("east_west", BuiltInParameter.BASEPOINT_EASTWEST_PARAM),
            ("elevation", BuiltInParameter.BASEPOINT_ELEVATION_PARAM),
            ("angle_to_true_north", BuiltInParameter.BASEPOINT_ANGLETON_PARAM),
        ]:
            row[key] = bip_str(bp, bip)
        result[label] = row
    result["active_project_location"] = doc.ActiveProjectLocation.Name
    return result


def export_links():
    rvt_types = []
    for lt in FilteredElementCollector(doc).OfClass(RevitLinkType):
        path = None
        try:
            path = ModelPathUtils.ConvertModelPathToUserVisiblePath(
                lt.GetExternalFileReference().GetAbsolutePath()
            )
        except Exception:
            pass
        rvt_types.append(
            {
                "name": ename(lt),
                "path": path,
                "loaded": RevitLinkType.IsLoaded(doc, lt.Id),
                "attachment": str(lt.AttachmentType),
                "workset": workset_name(lt),
            }
        )
    rvt_instances = [
        {
            "name": ename(li),
            "type": el_name(li.GetTypeId()),
            "pinned": li.Pinned,
            "workset": workset_name(li),
        }
        for li in FilteredElementCollector(doc).OfClass(RevitLinkInstance)
    ]
    cad = []
    for ii in FilteredElementCollector(doc).OfClass(ImportInstance):
        cad.append(
            {
                "name": el_name(ii.GetTypeId()),
                "is_linked": ii.IsLinked,
                "view_specific": ii.ViewSpecific,
                "owner_view": el_name(ii.OwnerViewId),
                "pinned": ii.Pinned,
                "workset": workset_name(ii),
                "id": idv(ii.Id),
            }
        )
    return {
        "rvt_link_types": rvt_types,
        "rvt_link_instances": rvt_instances,
        "cad_instances": cad,
    }


def export_categories():
    counts = {}
    for el in FilteredElementCollector(doc).WhereElementIsNotElementType():
        cat = el.Category
        if cat is None or cat.CategoryType != CategoryType.Model:
            continue
        counter_add(counts, cat.Name)
    return top(counts, limit=1000)


def type_row(element_type, instance_count):
    row = {
        "id": idv(element_type.Id),
        "category": element_type.Category.Name if element_type.Category else None,
        "family": element_type.FamilyName,
        "type": ename(element_type),
        "instances": instance_count,
        "type_mark": bip_str(element_type, BuiltInParameter.ALL_MODEL_TYPE_MARK),
        "keynote": bip_str(element_type, BuiltInParameter.KEYNOTE_PARAM),
        "description": bip_str(element_type, BuiltInParameter.ALL_MODEL_DESCRIPTION),
        "fire_rating": bip_str(element_type, BuiltInParameter.FIRE_RATING),
        "function": bip_str(element_type, BuiltInParameter.FUNCTION_PARAM),
    }
    if isinstance(element_type, HostObjAttributes):
        cs = element_type.GetCompoundStructure()
        row["compound_width_mm"] = to_mm(cs.GetWidth()) if cs is not None else None
        if cs is not None:
            first_core, last_core = (
                cs.GetFirstCoreLayerIndex(),
                cs.GetLastCoreLayerIndex(),
            )
            layers = []
            for i, layer in enumerate(cs.GetLayers()):
                material = doc.GetElement(layer.MaterialId)
                layers.append(
                    {
                        "index": i + 1,
                        "material": ename(material) if material is not None else None,
                        "material_id": idv(layer.MaterialId),
                        "width_mm": to_mm(layer.Width),
                        "function": str(layer.Function),
                        "core": first_core <= i <= last_core,
                    }
                )
            row["layers"] = layers
    if isinstance(element_type, WallType):
        row["wall_kind"] = str(element_type.Kind)
    if isinstance(element_type, FamilySymbol):
        row["in_place"] = element_type.Family.IsInPlace
        for key, bips in [
            (
                "width_mm",
                [
                    BuiltInParameter.DOOR_WIDTH,
                    BuiltInParameter.WINDOW_WIDTH,
                    BuiltInParameter.FAMILY_WIDTH_PARAM,
                ],
            ),
            (
                "height_mm",
                [
                    BuiltInParameter.DOOR_HEIGHT,
                    BuiltInParameter.WINDOW_HEIGHT,
                    BuiltInParameter.FAMILY_HEIGHT_PARAM,
                ],
            ),
        ]:
            for bip in bips:
                p = element_type.get_Parameter(bip)
                if p is not None and p.HasValue and p.StorageType == StorageType.Double:
                    row[key] = to_mm(p.AsDouble())
                    break
    railing_height = element_type.LookupParameter("Railing Height")
    if railing_height is not None and railing_height.StorageType == StorageType.Double:
        row["railing_height_mm"] = to_mm(railing_height.AsDouble())
    return row


def export_materials():
    rows = []
    for m in FilteredElementCollector(doc).OfClass(Material):
        rows.append(
            {
                "id": idv(m.Id),
                "name": ename(m),
                "class": m.MaterialClass,
                "category": m.MaterialCategory,
            }
        )
    return sorted(rows, key=lambda r: (r["name"] or ""))


def export_types():
    counts = {}
    for el in FilteredElementCollector(doc).WhereElementIsNotElementType():
        cat = el.Category
        if cat is None or cat.CategoryType != CategoryType.Model:
            continue
        type_id = el.GetTypeId()
        if type_id is not None and type_id != ElementId.InvalidElementId:
            counter_add(counts, idv(type_id))
    rows = []
    for element_type in FilteredElementCollector(doc).WhereElementIsElementType():
        cat = element_type.Category
        if cat is None or cat.CategoryType != CategoryType.Model:
            continue
        rows.append(type_row(element_type, counts.get(idv(element_type.Id), 0)))
    return sorted(
        rows, key=lambda r: (r["category"] or "", r["family"] or "", r["type"] or "")
    )


def instance_extras(group, el):
    extra = {}
    if group in ("MUR", "FAC"):
        extra["base_constraint"] = bip_str(el, BuiltInParameter.WALL_BASE_CONSTRAINT)
        extra["top_constraint"] = bip_str(el, BuiltInParameter.WALL_HEIGHT_TYPE)
    elif group == "PAV":
        try:
            extra["shape_edited"] = el.GetSlabShapeEditor().IsEnabled
        except Exception:
            extra["shape_edited"] = None
    elif group in ("POR", "FIN", "PFC"):
        host = getattr(el, "Host", None)
        if host is not None:
            host_type = doc.GetElement(host.GetTypeId())
            extra["host_id"] = idv(host.Id)
            extra["host_type"] = ename(host_type) if host_type is not None else None
            extra["host_type_mark"] = (
                bip_str(host_type, BuiltInParameter.ALL_MODEL_TYPE_MARK)
                if host_type is not None
                else None
            )
            if (
                isinstance(host_type, HostObjAttributes)
                and host_type.GetCompoundStructure() is not None
            ):
                extra["host_width_mm"] = to_mm(
                    host_type.GetCompoundStructure().GetWidth()
                )
    return extra


def export_pir_groups():
    result = []
    for code, bics, params in PIR_GROUPS:
        elements = collect_instances(bics)
        if code in ("MUR", "FAC"):
            want_curtain = code == "FAC"
            elements = [
                e
                for e in elements
                if isinstance(e, Wall)
                and (e.WallType.Kind == WallKind.Curtain) == want_curtain
            ]
        by_type, by_workset, by_phase_c, by_phase_d, by_level = {}, {}, {}, {}, {}
        stats = []
        for name, pir_level in params:
            stats.append(
                {
                    "param": name,
                    "pir_level": pir_level,
                    "found_on": {},
                    "states": {},
                    "values": {},
                    "not_valued_ids": [],
                }
            )
        extras = []
        wbs_by_type = {}  # "category|type|type mark" -> {"L7|L8": count}: to check level 8 per type
        for el in elements:
            counter_add(by_type, ename(el))
            if R.WBS_LEVELS[6] in [n for n, _ in params]:
                el_type = doc.GetElement(el.GetTypeId())
                key = "%s|%s|%s" % (
                    el.Category.Name if el.Category else "",
                    ename(el_type) if el_type is not None else "",
                    bip_str(el_type, BuiltInParameter.ALL_MODEL_TYPE_MARK) if el_type is not None else "",
                )
                pair = "|".join((read_param(el, n)[1] or "") for n in R.WBS_LEVELS[6:8])
                counter_add(wbs_by_type.setdefault(key, {}), pair)
            counter_add(by_workset, workset_name(el))
            counter_add(by_phase_c, bip_str(el, BuiltInParameter.PHASE_CREATED))
            counter_add(by_phase_d, bip_str(el, BuiltInParameter.PHASE_DEMOLISHED))
            level_id = getattr(el, "LevelId", None)
            counter_add(by_level, el_name(level_id) if level_id is not None else None)
            for stat in stats:
                found_on, state, value = read_param_any(el, stat["param"])
                counter_add(stat["found_on"], found_on)
                counter_add(stat["states"], state)
                if state == "value" or state == "nd":
                    counter_add(stat["values"], value)
                elif len(stat["not_valued_ids"]) < MAX_IDS:
                    stat["not_valued_ids"].append(idv(el.Id))
            extra = instance_extras(code, el)
            if extra:
                extra["id"] = idv(el.Id)
                extra["type"] = ename(el)
                extras.append(extra)
        for stat in stats:
            stat["values"] = top(stat["values"])
        result.append(
            {
                "group": code,
                "categories": [str(b) for b in bics],
                "count": len(elements),
                "by_type": top(by_type, limit=1000),
                "by_workset": top(by_workset),
                "by_phase_created": top(by_phase_c),
                "by_phase_demolished": top(by_phase_d),
                "by_level": top(by_level),
                "parameters": stats,
                "instance_extras": extras,
                "wbs_by_type": wbs_by_type,
            }
        )
    return result


def export_rooms():
    options = SpatialElementBoundaryOptions()
    rows = []
    numbers = {}
    for room in (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Rooms)
        .WhereElementIsNotElementType()
    ):
        number = bip_str(room, BuiltInParameter.ROOM_NUMBER)
        counter_add(numbers, number)
        placed = room.Location is not None
        segments = room.GetBoundarySegments(options) if placed else None
        rows.append(
            {
                "id": idv(room.Id),
                "number": number,
                "name": bip_str(room, BuiltInParameter.ROOM_NAME),
                "level": el_name(room.LevelId),
                "placed": placed,
                "area_m2": round(room.Area * 0.09290304, 3),
                "boundary_loops": segments.Count if segments is not None else 0,
                "upper_limit": bip_str(room, BuiltInParameter.ROOM_UPPER_LEVEL),
                "upper_offset": bip_str(room, BuiltInParameter.ROOM_UPPER_OFFSET),
                "phase": (
                    bip_str(room, BuiltInParameter.ROOM_PHASE)
                    if hasattr(BuiltInParameter, "ROOM_PHASE")
                    else None
                ),
            }
        )
    return {
        "rooms": rows,
        "duplicate_numbers": [[k, v] for k, v in numbers.items() if v > 1],
    }


def export_warnings():
    groups = {}
    for fm in doc.GetWarnings():
        key = fm.GetDescriptionText()
        entry = groups.setdefault(
            key,
            {
                "description": key,
                "severity": str(fm.GetSeverity()),
                "count": 0,
                "element_ids": [],
            },
        )
        entry["count"] += 1
        if len(entry["element_ids"]) < 200:
            entry["element_ids"].append([idv(i) for i in fm.GetFailingElements()])
    return sorted(groups.values(), key=lambda g: -g["count"])


def export_views():
    on_sheet = {}
    for vp in FilteredElementCollector(doc).OfClass(Viewport):
        on_sheet[idv(vp.ViewId)] = el_name(vp.SheetId)
    views = []
    for v in FilteredElementCollector(doc).OfClass(View):
        if isinstance(v, ViewSheet):
            continue
        row = {
            "id": idv(v.Id),
            "name": v.Name,
            "view_type": str(v.ViewType),
            "is_template": v.IsTemplate,
            "template": el_name(v.ViewTemplateId),
            "on_sheet": on_sheet.get(idv(v.Id)),
            "phase": bip_str(v, BuiltInParameter.VIEW_PHASE),
            "phase_filter": bip_str(v, BuiltInParameter.VIEW_PHASE_FILTER),
        }
        for name in (
            "SGAS_View Status",
            "SGAS_View Topic",
            "SGAS_View Type",
            "VIS_TE_Status della vista",
            "VIS_TE_Argomento di vista",
            "VI_TE_Argomento della vista",
            "VI_TE_Tema della vista",
        ):
            state, value = read_param(v, name)
            if state != "missing":
                row[name] = value
        views.append(row)
    sheets = [
        {"id": idv(s.Id), "number": s.SheetNumber, "name": s.Name}
        for s in FilteredElementCollector(doc).OfClass(ViewSheet)
    ]
    return {
        "views": sorted(views, key=lambda r: r["name"]),
        "sheets": sorted(sheets, key=lambda r: r["number"]),
        "schedules": export_schedule_fields(),
        "filters": export_filter_parameters(),
    }


def param_label(param_id):
    """Project/shared parameter name for positive ids, built-in parameter name otherwise."""
    if param_id is None or param_id == ElementId.InvalidElementId:
        return None
    if idv(param_id) > 0:
        return el_name(param_id)
    try:
        return str(System.Enum.ToObject(BuiltInParameter, idv(param_id)))
    except Exception:
        return str(idv(param_id))


def export_schedule_fields():
    """Parameters used by the schedule fields: what a parameter deletion would break."""
    result = []
    for vs in FilteredElementCollector(doc).OfClass(ViewSchedule):
        if vs.IsTemplate or vs.IsTitleblockRevisionSchedule:
            continue
        try:
            d = vs.Definition
            fields = []
            for i in range(d.GetFieldCount()):
                f = d.GetField(i)
                fields.append(
                    {"heading": f.GetName(), "param": param_label(f.ParameterId)}
                )
            cat = (
                Category.GetCategory(doc, d.CategoryId)
                if d.CategoryId != ElementId.InvalidElementId
                else None
            )
            result.append(
                {
                    "id": idv(vs.Id),
                    "name": vs.Name,
                    "category": cat.Name if cat else None,
                    "fields": fields,
                }
            )
        except Exception as ex:
            result.append({"id": idv(vs.Id), "name": vs.Name, "error": str(ex)})
    return sorted(result, key=lambda r: r["name"])


def export_filter_parameters():
    """View filters and the parameters of their rules."""
    result = []
    for pf in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
        result.append(
            {
                "id": idv(pf.Id),
                "name": pf.Name,
                "categories": sorted(el_name_cat(c) for c in pf.GetCategories()),
                "params": sorted(
                    param_label(p) or "" for p in pf.GetElementFilterParameters()
                ),
            }
        )
    return sorted(result, key=lambda r: r["name"])


def el_name_cat(category_id):
    cat = Category.GetCategory(doc, category_id)
    return cat.Name if cat is not None else str(idv(category_id))


def export_cleanup():
    in_place = []
    for fam in FilteredElementCollector(doc).OfClass(Family):
        if fam.IsInPlace:
            in_place.append(
                {
                    "name": fam.Name,
                    "category": fam.FamilyCategory.Name if fam.FamilyCategory else None,
                }
            )
    groups = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_IOSModelGroups)
        .WhereElementIsNotElementType()
        .GetElementCount()
    )
    unused = {}
    for element_id in doc.GetUnusedElements(HashSet[ElementId]()):
        element = doc.GetElement(element_id)
        cat = (
            element.Category.Name
            if element is not None and element.Category is not None
            else "(no category)"
        )
        counter_add(unused, cat)
    return {
        "in_place_families": in_place,
        "model_group_instances": groups,
        "unused_elements_by_category": top(unused, limit=1000),
    }


# ---------------------------------------------------------------- main

folder = forms.pick_folder(title="Select the output folder for the JSON export")
if not folder:
    forms.alert("No folder selected.", exitscript=True)

data = {}
sections = [
    ("model_info", export_model_info),
    ("project_parameters", export_project_parameters),
    ("parameter_usage", export_parameter_usage),
    ("phases", export_phases),
    ("worksets", export_worksets),
    ("datums", export_datums),
    ("base_points", export_base_points),
    ("links", export_links),
    ("categories", export_categories),
    ("materials", export_materials),
    ("types", export_types),
    ("pir_groups", export_pir_groups),
    ("rooms", export_rooms),
    ("warnings", export_warnings),
    ("views", export_views),
    ("cleanup", export_cleanup),
]

output.print_md("# BMC - Model Data Export")
output.print_md("Model: **{}**".format(doc.Title))
for key, func in sections:
    data[key] = run(key, func)
    output.print_md("- {} ({} s)".format(key, timings[-1][1]))

data["errors"] = errors
data["timings_s"] = timings

stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
file_path = os.path.join(folder, "BMC_ModelData_{}_{}.json".format(doc.Title, stamp))
# ensure_ascii=False: IronPython json re-decodes str as utf-8 when ensure_ascii=True and fails on chars like the degree sign
with codecs.open(file_path, "w", encoding="utf-8") as f:
    f.write(json.dumps(data, indent=1, ensure_ascii=False, default=str))

output.print_md("## Export completed")
output.print_md("File: `{}`".format(file_path))
if errors:
    output.print_md("## Sections with errors: {}".format(len(errors)))
    for err in errors:
        output.print_md("- **{}**: {}".format(err["section"], err["error"]))
else:
    output.print_md("No errors.")
