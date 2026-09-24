"""Merge every modelling workset into AR_ARCHITETTURA and put the areas on AR_AREE; coordination worksets ("_...") are kept"""

__title__ = "Workset\nunico"
__author__ = "Luca Rosati"
__guida__ = 'Unisce i workset di modellazione in AR_ARCHITETTURA, sposta le aree su AR_AREE e lascia invariati i workset di coordinamento (_...).'

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# PGI R06 3.11.8 + decision 23/09/2026: coordination worksets "_...", one modelling workset
# AR_ARCHITETTURA (the old "AR" is renamed), AR_AREE for areas.
# DeleteWorkset with MoveElementsToWorkset moves the elements and removes the workset in one call;
# when Revit refuses the deletion the elements are moved one by one and the empty workset is left.
# One Transaction (Ctrl+Z undoes everything) + JSON log.

import clr

clr.AddReference("System.Core")
from System.Collections.Generic import HashSet

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script

import bmc_rules as R
from bmc_utils import commit, instances, save_log, workset_name

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

OLD_MODEL_NAMES = [
    "AR"
]  # decision 21/09/2026 name, renamed to R.WS_MODEL on 23/09/2026
AREA_CATS = [BuiltInCategory.OST_Areas, BuiltInCategory.OST_AreaSchemeLines]

if not doc.IsWorkshared:
    forms.alert(
        "Il modello non e' condiviso (worksharing): nessun workset da gestire.",
        exitscript=True,
    )


def user_worksets():
    return list(FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset))


def by_name(name):
    found = [w for w in user_worksets() if w.Name == name]
    return found[0] if found else None


def count(ws):
    return (
        FilteredElementCollector(doc)
        .WherePasses(ElementWorksetFilter(ws.Id))
        .GetElementCount()
    )


def why_not_deletable(ws):
    """Revit only says 'cannot delete': report editability, owner and elements borrowed by others."""
    ws = doc.GetWorksetTable().GetWorkset(ws.Id)
    borrowed = {}
    for el in FilteredElementCollector(doc).WherePasses(ElementWorksetFilter(ws.Id)):
        if (
            WorksharingUtils.GetCheckoutStatus(doc, el.Id)
            == CheckoutStatus.OwnedByOtherUser
        ):
            owner = WorksharingUtils.GetWorksharingTooltipInfo(doc, el.Id).Owner
            borrowed[owner] = borrowed.get(owner, 0) + 1
    return (
        "non eliminabile: modificabile=%s, proprietario='%s', elementi presi da altri=%s"
        % (
            ws.IsEditable,
            ws.Owner,
            borrowed or "nessuno",
        )
    )


def move(elements, target):
    """Move elements to target via their Workset parameter. Returns (moved, {reason: [ids]})."""
    moved, not_moved = 0, {}
    for el in elements:
        cat = el.Category.Name if el.Category is not None else el.GetType().Name
        p = el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        try:
            if p is None or p.IsReadOnly:
                raise Exception("workset in sola lettura")
            p.Set(target.Id.IntegerValue)
            moved += 1
        except Exception as ex:
            not_moved.setdefault("%s: %s" % (cat, ex), []).append(el.Id.ToString())
    return moved, not_moved


# ---------------------------------------------------------------- plan (read only)

target = by_name(R.WS_MODEL)
rename = (
    None
    if target
    else next((w for w in user_worksets() if w.Name in OLD_MODEL_NAMES), None)
)
areas_ws = by_name(R.WS_AREAS)
special = [R.WS_MODEL, R.WS_AREAS] + ([rename.Name] if rename else [])
merge = [
    w for w in user_worksets() if w.Name not in special and not w.Name.startswith("_")
]
keep = [w for w in user_worksets() if w.Name.startswith("_")]
areas = [a for a in instances(doc, AREA_CATS) if workset_name(doc, a) != R.WS_AREAS]

output.print_md("# BMC - Workset unico (anteprima)")
output.print_md("Modello: **%s**. Nessuna modifica e' stata ancora fatta." % doc.Title)
rows = [[w.Name, count(w), "unito in %s ed eliminato" % R.WS_MODEL] for w in merge]
rows += [[w.Name, count(w), "mantenuto (coordinamento)"] for w in keep]
if target:
    rows.append([R.WS_MODEL, count(target), "esistente"])
elif rename:
    rows.append([rename.Name, count(rename), "rinominato in %s" % R.WS_MODEL])
else:
    rows.append([R.WS_MODEL, 0, "da creare"])
rows.append(
    [
        R.WS_AREAS,
        count(areas_ws) if areas_ws else 0,
        "esistente" if areas_ws else "da creare",
    ]
)
output.print_table(table_data=rows, columns=["Workset", "Elementi", "Azione"])
output.print_md(
    "Aree e linee di confine da spostare su `%s`: **%d**" % (R.WS_AREAS, len(areas))
)

if not merge and target and not areas:
    forms.alert("Workset gia' conformi: niente da fare.", exitscript=True)
if not forms.alert(
    "Unire %d workset in '%s'%s e spostare %d aree su '%s'?\n"
    "I workset che iniziano con '_' restano invariati. L'operazione si annulla con Ctrl+Z."
    % (
        len(merge),
        R.WS_MODEL,
        " (rinominando '%s')" % rename.Name if rename else "",
        len(areas),
        R.WS_AREAS,
    ),
    yes=True,
    no=True,
):
    script.exit()

# ---------------------------------------------------------------- apply

log = {"model": doc.Title, "merged": [], "failed": [], "renamed": None, "areas": None}

# worksets must be editable to be deleted or renamed: check them out when the model has a central file
try:
    ids = HashSet[WorksetId]()
    for w in merge + [x for x in (target, rename, areas_ws) if x]:
        ids.Add(w.Id)
    WorksharingUtils.CheckoutWorksets(doc, ids)
except Exception as ex:
    log["checkout_error"] = str(ex)

t = Transaction(doc, "BMC - Workset unico")
t.Start()
try:
    if target:
        model_ws = target
    elif rename:
        WorksetTable.RenameWorkset(doc, rename.Id, R.WS_MODEL)
        model_ws = doc.GetWorksetTable().GetWorkset(rename.Id)
        log["renamed"] = [rename.Name, R.WS_MODEL]
    else:
        model_ws = Workset.Create(doc, R.WS_MODEL)
    try:
        # the active workset cannot be deleted: make the modelling workset active first
        doc.GetWorksetTable().SetActiveWorksetId(model_ws.Id)
    except Exception as ex:
        log["active_error"] = str(ex)
    for w in merge:
        entry = {"workset": w.Name, "elements": count(w)}
        settings = DeleteWorksetSettings(
            DeleteWorksetOption.MoveElementsToWorkset, model_ws.Id
        )
        try:
            if not WorksetTable.CanDeleteWorkset(doc, w.Id, settings):
                raise Exception(why_not_deletable(w))
            WorksetTable.DeleteWorkset(doc, w.Id, settings)
            log["merged"].append(entry)
        except Exception as ex:
            entry["error"] = str(ex)
            entry["moved"], entry["not_moved"] = move(
                list(
                    FilteredElementCollector(doc).WherePasses(
                        ElementWorksetFilter(w.Id)
                    )
                ),
                model_ws,
            )
            log["failed"].append(entry)
    area_target = areas_ws or Workset.Create(doc, R.WS_AREAS)
    # areas collected again: the merge may have moved them to the modelling workset
    moved, not_moved = move(
        [a for a in instances(doc, AREA_CATS) if workset_name(doc, a) != R.WS_AREAS],
        area_target,
    )
    log["areas"] = {"moved": moved, "not_moved": not_moved}
except Exception as ex:
    t.RollBack()
    forms.alert(
        "Errore imprevisto, nessuna modifica applicata:\n%s" % ex, exitscript=True
    )
else:
    log["commit_status"] = commit(t)

log_path = save_log("BMC_Workset", log)

output.print_md("---")
if log["commit_status"] != "Committed":
    output.print_md(
        "# ATTENZIONE: transazione %s, il modello NON e' stato modificato"
        % log["commit_status"]
    )
output.print_md(
    "# Workset uniti: %d - non eliminati: %d" % (len(log["merged"]), len(log["failed"]))
)
if log["renamed"]:
    output.print_md("Rinominato `%s` in `%s`" % tuple(log["renamed"]))
if log["failed"]:
    output.print_md(
        "## Workset non eliminati: elementi spostati uno per uno su %s" % R.WS_MODEL
    )
    output.print_table(
        table_data=[
            [f["workset"], f["elements"], f["moved"], f["error"]] for f in log["failed"]
        ],
        columns=["Workset", "Elementi", "Spostati", "Motivo mancata eliminazione"],
    )
not_moved = [
    (f["workset"], r, i) for f in log["failed"] for r, i in f["not_moved"].items()
]
not_moved += [(R.WS_AREAS, r, i) for r, i in log["areas"]["not_moved"].items()]
for ws_name, reason, ids in sorted(not_moved):
    output.print_md(
        "- `%s` non spostati (%d) - %s %s"
        % (
            ws_name,
            len(ids),
            reason,
            output.linkify([ElementId(int(i)) for i in ids[:200]]),
        )
    )
output.print_md("Aree spostate su `%s`: %d" % (R.WS_AREAS, log["areas"]["moved"]))
if "checkout_error" in log:
    output.print_md("Checkout workset non riuscito: %s" % log["checkout_error"])
output.print_md("## Situazione finale")
output.print_table(
    table_data=[[w.Name, count(w)] for w in user_worksets()],
    columns=["Workset", "Elementi"],
)
output.print_md("Log: `%s`" % log_path)
