"""Helpers shared by the BMC.tab scripts (read parameters, collect elements, logs).

Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
"""

import codecs
import datetime
import json
import os

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Element,
    ElementId,
    FilteredElementCollector,
    StorageType,
    UnitTypeId,
    UnitUtils,
    Wall,
    WallKind,
)

import bmc_rules as R


def ename(element):
    """Name of any element. ElementType redeclares Name as set-only: read it through Element."""
    return Element.Name.GetValue(element)


def idv(element_id):
    """Integer value of an ElementId (Revit 2024 Value, older IntegerValue)."""
    try:
        return int(element_id.Value)
    except AttributeError:
        return element_id.IntegerValue


def to_mm(value_ft):
    return UnitUtils.ConvertFromInternalUnits(value_ft, UnitTypeId.Millimeters)


def to_m(value_ft):
    return UnitUtils.ConvertFromInternalUnits(value_ft, UnitTypeId.Meters)


def get_param(element, name):
    """Built-in parameters by BuiltInParameter (language independent), shared ones by name."""
    if name in R.BUILTIN:
        return element.get_Parameter(R.BUILTIN[name])
    return element.LookupParameter(name)


def param_state(element, name):
    """(state, value). state: missing / empty / nd / value. Never-set and blank text count as empty."""
    p = get_param(element, name)
    if p is None:
        return "missing", None
    if not p.HasValue:
        return "empty", None
    if p.StorageType == StorageType.String:
        raw = (p.AsString() or "").strip()
        if not raw:
            return "empty", raw
        if raw == R.ND:
            return "nd", raw
        return "value", raw
    if p.StorageType == StorageType.Integer:
        return "value", p.AsInteger()
    if p.StorageType == StorageType.ElementId:
        return "value", p.AsValueString()
    return "value", p.AsValueString()


def type_of(doc, element):
    type_id = element.GetTypeId()
    if type_id is None or type_id == ElementId.InvalidElementId:
        return None
    return doc.GetElement(type_id)


def param_state_any(doc, element, name, level="I"):
    """Returns (target element, state, value). level "I": instance first, then its type;
    level "T" (PIR "Definisce il tipo"): type first. Built-ins such as Description also exist, empty,
    on door and furniture instances: reading them there gave false "not compiled" results.
    """
    element_type = type_of(doc, element)
    order = [element_type, element] if level == "T" else [element, element_type]
    for target in order:
        if target is None:
            continue
        state, value = param_state(target, name)
        if state != "missing":
            return target, state, value
    return element, "missing", None


def instances(doc, bics):
    result = []
    for bic in bics:
        result.extend(
            FilteredElementCollector(doc)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
            .ToElements()
        )
    return result


def used_types(doc, bics):
    """[(type element, instance count)] for the types used in the given categories."""
    counts = {}
    for el in instances(doc, bics):
        type_id = el.GetTypeId()
        if type_id != ElementId.InvalidElementId:
            counts.setdefault(idv(type_id), [type_id, 0])[1] += 1
    result = []
    for type_id, n in counts.values():
        element_type = doc.GetElement(type_id)
        if element_type is not None:
            result.append((element_type, n))
    return result


def workset_name(doc, element):
    try:
        return doc.GetWorksetTable().GetWorkset(element.WorksetId).Name
    except Exception:
        return None


def type_mark(element_type):
    p = element_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK)
    return (p.AsString() or "").strip() if p is not None else ""


def layer_codes(doc, element_type):
    """Material codes (XXX.NN) of the compound structure layers of a type, [] when it has no layers."""
    get_cs = getattr(element_type, "GetCompoundStructure", None)
    cs = get_cs() if get_cs else None
    if cs is None:
        return []
    codes = []
    for layer in cs.GetLayers():
        material = doc.GetElement(layer.MaterialId)
        m = R.RE_MATERIAL.match(ename(material)) if material is not None else None
        codes.append(m.group(1) if m else "?")
    return codes


def in_curtain_wall(element):
    """True when the element (door, window) is hosted by a curtain wall."""
    host = getattr(element, "Host", None)
    return isinstance(host, Wall) and host.WallType.Kind == WallKind.Curtain


def host_codes(doc, element):
    """(Type Mark of the host, Type Mark_thickness of the host) of a hosted element (PIR SPE HOST).
    ("", None) without host. Hosts without layers (curtain walls, e.g. AR_MUR_CV1.01G_CUR_150) take the
    thickness from their type name; None when neither exists. The Type Mark is returned as it is:
    callers check it with R.RE_TYPE_MARK (OLD_ types have "OLD")."""
    host = getattr(element, "Host", None)
    host_type = doc.GetElement(host.GetTypeId()) if host is not None else None
    if host_type is None:
        return "", None
    tm = type_mark(host_type)
    get_cs = getattr(host_type, "GetCompoundStructure", None)
    cs = get_cs() if get_cs else None
    if cs is not None:
        width = to_mm(cs.GetWidth())
    else:
        m = R.RE_STRAT_TYPE.match(ename(host_type))
        width = float(m.group(7)) if m else None
    if not tm or width is None:
        return tm, None
    return tm, R.host_thickness_code(tm, width)


def level_name(doc, element):
    """Reference level of an element (LevelId, then the usual level parameters)."""
    level_id = getattr(element, "LevelId", None)
    if level_id is None or level_id == ElementId.InvalidElementId:
        for bip in (
            BuiltInParameter.FAMILY_LEVEL_PARAM,
            BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
            BuiltInParameter.WALL_BASE_CONSTRAINT,
            BuiltInParameter.ROOM_LEVEL_ID,
            BuiltInParameter.LEVEL_PARAM,
            BuiltInParameter.SCHEDULE_LEVEL_PARAM,
        ):
            p = element.get_Parameter(bip)
            if p is not None and p.StorageType == StorageType.ElementId:
                level_id = p.AsElementId()
                if level_id != ElementId.InvalidElementId:
                    break
    if level_id is None or level_id == ElementId.InvalidElementId:
        return None
    level = doc.GetElement(level_id)
    return ename(level) if level is not None else None


PREVIEW, APPLY = "Solo anteprima", "Anteprima e applica"


def ask_preview_only():
    """Ask the mode BEFORE computing the preview. A modal WPF dialog (SelectFromList) disables every window of
    the thread, the pyRevit output window too: the preview cannot be scrolled while the groups are chosen.
    So the preview is read in a first run and applied in a second one. Returns True for preview only.
    """
    from pyrevit import forms, script

    choice = forms.CommandSwitchWindow.show(
        [PREVIEW, APPLY],
        # title is required: without it pyRevit master looks up a localized title and its except clause
        # references wpf.ResourceReferenceKeyNotFoundException, missing in IronPython (AttributeError)
        title="BMC - anteprima o applica",
        message="Leggi prima l'anteprima, poi rilancia il comando e scegli '%s'"
        % APPLY,
    )
    if not choice:
        script.exit()
    return choice == PREVIEW


def end_preview(output, preview_only):
    """Stop after the preview in preview-only mode: the output window stays free to scroll."""
    if not preview_only:
        return
    from pyrevit import script

    output.print_md("---")
    output.print_md(
        "**Solo anteprima: il modello non e' stato modificato.** "
        "Per applicare rilancia il comando e scegli *%s*." % APPLY
    )
    script.exit()


def save_log(prefix, data):
    """Write data as JSON in Documents\\<prefix>_<stamp>.json and return the path."""
    path = os.path.join(
        os.path.expanduser("~"),
        "Documents",
        "%s_%s.json" % (prefix, datetime.datetime.now().strftime("%Y%m%d_%H%M%S")),
    )
    # default=str: ElementId / Int64 values are not JSON serializable in IronPython
    text = json.dumps(data, indent=1, ensure_ascii=False, default=str)
    with codecs.open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path


def commit(transaction):
    """Commit and return the status name; RolledBack means Revit showed an error that was cancelled."""
    return str(transaction.Commit())
