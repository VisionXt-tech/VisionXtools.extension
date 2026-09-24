"""Re-issue the project parameters as the R06 shared parameter files and the PIR require, keeping the values: backup -> delete/rebind/insert -> restore"""

__title__ = "Riemetti\nparametri R06"
__author__ = "Luca Rosati"
__guida__ = 'Riallinea i parametri condivisi ai TXT R06 e al PIR (categorie, tipo/istanza, GUID), elimina i legacy e ricompila i valori, abachi compresi. Backup JSON prima di ogni modifica.'

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# Sources: TXT R06 in BMC.tab/lib/data/shared_params (AR, GE, SPACE, BROWSER, TAVOLE, Supporto; not MEP/ST),
# categories and type/instance from bmc_rules.binding_targets() (PIR_ARC, PIR_SUP e AREE, Viste, Tavole).
# Actions per parameter:
#   aggiungi    missing -> inserted from its TXT
#   riassocia   same GUID, wrong categories or type/instance -> BindingMap.ReInsert (the parameter element is
#               kept, so schedules, filters and browser organisation keep working)
#   sostituisci wrong GUID -> the old parameter is deleted and the TXT one inserted
#   elimina     duplicated copies with a wrong GUID and legacy parameters (after migrating their useful values)
# Flow: plan (read only) -> preview -> backup JSON of every affected value (written BEFORE any change) ->
# one Transaction: delete, rebind, insert, restore values, migrations -> report + log. Ctrl+Z undoes all.

import clr
from System import Guid
import os
from collections import Counter

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script

import bmc_rules as R
from bmc_utils import commit, idv, save_log

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
output = script.get_output()

MAX_ROWS = 80


def cat_names(bics):
    names, refused = [], []
    for bic in bics:
        cat = Category.GetCategory(doc, bic)
        if cat is None or not cat.AllowsBoundParameters:
            refused.append(str(bic))
        else:
            names.append(cat.Name)
    return names, refused


def category_set(bics):
    cs = app.Create.NewCategorySet()
    for bic in bics:
        cat = Category.GetCategory(doc, bic)
        if cat is not None and cat.AllowsBoundParameters:
            cs.Insert(cat)
    return cs


def make_binding(level, bics):
    cs = category_set(bics)
    return (
        app.Create.NewTypeBinding(cs)
        if level == "T"
        else app.Create.NewInstanceBinding(cs)
    )


def read_value(p):
    """(storage, value) of a parameter with a value, else None."""
    if p is None or not p.HasValue:
        return None
    st = p.StorageType
    if st == StorageType.String:
        text = p.AsString()
        return ("s", text) if text is not None and text.strip() else None
    if st == StorageType.Integer:
        return ("i", p.AsInteger())
    if st == StorageType.Double:
        return ("d", p.AsDouble())
    return None  # ElementId values are not used by these parameters


def is_empty(p):
    if not p.HasValue:
        return True
    return p.StorageType == StorageType.String and not (p.AsString() or "").strip()


def write_value(p, storage, value, conversion=None):
    if conversion == "area_to_m2" and storage == "d":
        value, storage = (
            UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.SquareMeters),
            "d",
        )
    st = p.StorageType
    if st == StorageType.String:
        p.Set(value if storage == "s" else str(value))
    elif st == StorageType.Integer:
        p.Set(int(value) if storage != "s" else int(float(value.replace(",", "."))))
    elif st == StorageType.Double:
        p.Set(float(value) if storage != "s" else float(value.replace(",", ".")))
    else:
        raise Exception("tipo di parametro non gestito")


# ---------------------------------------------------------------- current state (read only)

current = {}  # name -> [entry]
it = doc.ParameterBindings.ForwardIterator()
it.Reset()
while it.MoveNext():
    definition, binding = it.Key, it.Current
    spe = doc.GetElement(definition.Id)
    current.setdefault(definition.Name, []).append(
        {
            "definition": definition,
            "name": definition.Name,
            # read now: the definition is no longer valid once its parameter element is deleted
            "group": definition.GetGroupTypeId(),
            "pid": definition.Id,
            "guid": (
                str(spe.GuidValue).lower()
                if isinstance(spe, SharedParameterElement)
                else None
            ),
            "level": "T" if isinstance(binding, TypeBinding) else "I",
            "cats": sorted(c.Name for c in binding.Categories),
            "cat_ids": [idv(c.Id) for c in binding.Categories],
        }
    )

# ---------------------------------------------------------------- plan (read only)

targets = R.binding_targets()
plan = (
    []
)  # dict(name, action, level, bics, file, delete=[entries], keep=entry|None, note)
for name, (level, bics) in sorted(targets.items()):
    ref = R.SHARED_PARAMS[name]
    wanted, refused = cat_names(bics)
    entries = current.get(name, [])
    keep = [e for e in entries if e["guid"] == ref["guid"].lower()]
    wrong = [e for e in entries if e not in keep]
    item = {
        "name": name,
        "file": ref["file"],
        "level": level,
        "bics": bics,
        "wanted": sorted(wanted),
        "refused": refused,
        "target_cat_ids": [
            idv(Category.GetCategory(doc, b).Id)
            for b in bics
            if Category.GetCategory(doc, b)
        ],
        "keep": keep[0] if keep else None,
        "delete": wrong + keep[1:],
    }
    if not entries:
        item["action"] = "aggiungi"
    elif not keep:
        item["action"] = "sostituisci"
    elif keep[0]["level"] != level or keep[0]["cats"] != sorted(wanted):
        item["action"] = "riassocia"
    else:
        item["action"] = "ok"
    if item["action"] == "ok" and item["delete"]:
        item["action"] = "elimina copie"
    plan.append(item)

legacy = []
for name in R.LEGACY_PARAMS:
    for entry in current.get(name, []):
        legacy.append({"name": name, "action": "elimina legacy", "delete": [entry]})

unmanaged = sorted(n for n in current if n not in targets and n not in R.LEGACY_PARAMS)


def dependencies(pids):
    """Schedules and view filters that use the parameter elements about to be deleted."""
    pids = set(idv(p) for p in pids)
    schedules, filters = {}, {}
    for vs in FilteredElementCollector(doc).OfClass(ViewSchedule):
        try:
            d = vs.Definition
            for i in range(d.GetFieldCount()):
                pid = idv(d.GetField(i).ParameterId)
                if pid in pids:
                    schedules.setdefault(pid, []).append(vs.Name)
        except Exception:
            pass
    for pf in FilteredElementCollector(doc).OfClass(ParameterFilterElement):
        for p in pf.GetElementFilterParameters():
            if idv(p) in pids:
                filters.setdefault(idv(p), []).append(pf.Name)
    return schedules, filters


to_delete = [e for item in plan + legacy for e in item["delete"]]
dep_schedules, dep_filters = dependencies([e["pid"] for e in to_delete])

# ---------------------------------------------------------------- preview

ACTIONS = ["aggiungi", "riassocia", "sostituisci", "elimina copie", "elimina legacy"]
by_action = dict(
    (a, [i for i in plan + legacy if i["action"] == a]) for a in ACTIONS + ["ok"]
)

output.print_md("# BMC - Riemissione parametri R06 (anteprima)")
output.print_md(
    "Modello: **%s**. Nessuna modifica e' stata ancora fatta. TXT: %s."
    % (doc.Title, ", ".join(sorted(R.SHARED_FILES)))
)
output.print_table(
    table_data=[[a, len(by_action[a])] for a in ACTIONS + ["ok"]],
    columns=["Azione", "Parametri"],
)
rows = []
for item in [i for i in plan if i["action"] != "ok"]:
    now = item["keep"] or (item["delete"][0] if item["delete"] else None)
    rows.append(
        [
            item["name"],
            item["file"],
            item["action"],
            "%s %s" % (now["level"], ", ".join(now["cats"])[:90]) if now else "-",
            "%s %s" % (item["level"], ", ".join(item["wanted"])[:90]),
        ]
    )
if rows:
    output.print_table(
        table_data=rows, columns=["Parametro", "TXT", "Azione", "Ora", "PIR R06"]
    )
if legacy:
    output.print_md("## Parametri legacy da eliminare")
    output.print_table(
        table_data=[
            [
                i["name"],
                ", ".join(i["delete"][0]["cats"])[:80],
                ", ".join(
                    "%s -> %s" % (m[0], m[1]) for m in R.MIGRATIONS if m[0] == i["name"]
                )
                or "-",
            ]
            for i in legacy
        ],
        columns=["Parametro", "Categorie", "Migrazione valori (solo dove vuoto)"],
    )
if dep_schedules or dep_filters:
    output.print_md("## ATTENZIONE: abachi e filtri che perdono un campo")
    names = dict((idv(e["pid"]), e["name"]) for e in to_delete)
    for pid, views in sorted(dep_schedules.items()):
        output.print_md(
            "- `%s` negli abachi: %s" % (names[pid], ", ".join(sorted(set(views))))
        )
    for pid, pfs in sorted(dep_filters.items()):
        output.print_md(
            "- `%s` nei filtri: %s" % (names[pid], ", ".join(sorted(set(pfs))))
        )
refused = [(i["name"], r) for i in plan for r in i["refused"]]
if refused:
    output.print_md(
        "Categorie che non accettano parametri (ignorate): %s"
        % ", ".join("%s/%s" % x for x in refused)
    )
if unmanaged:
    output.print_md(
        "Parametri non gestiti (non nei TXT, non legacy): %s" % ", ".join(unmanaged)
    )

groups = ["%s (%d)" % (a, len(by_action[a])) for a in ACTIONS if by_action[a]]
if not groups:
    forms.alert("Parametri gia' conformi ai TXT R06: niente da fare.", exitscript=True)
chosen = forms.SelectFromList.show(
    groups,
    title="Azioni da applicare (vedi anteprima)",
    multiselect=True,
    button_name="Applica",
)
if not chosen:
    forms.alert(
        "Nessuna azione selezionata: il modello non e' stato modificato.",
        exitscript=True,
    )
chosen = [c.rsplit(" (", 1)[0] for c in chosen]
selected = [i for i in plan + legacy if i["action"] in chosen]

# schedule columns that use a copy about to be deleted of a parameter that stays (replaced or duplicated):
# recorded now and re-inserted at the same position with the final parameter after the changes
deleted_kept_names = dict(
    (idv(e["pid"]), i["name"])
    for i in selected
    if i["action"] != "elimina legacy"
    for e in i["delete"]
)
field_moves = []  # (schedule id, index, field type, heading, hidden, parameter name)
for vs in FilteredElementCollector(doc).OfClass(ViewSchedule):
    try:
        d = vs.Definition
        for i in range(d.GetFieldCount()):
            f = d.GetField(i)
            name = deleted_kept_names.get(idv(f.ParameterId))
            if name:
                field_moves.append(
                    (vs.Id, i, f.FieldType, f.ColumnHeading, f.IsHidden, name)
                )
    except Exception:
        pass

# ---------------------------------------------------------------- backup (read only, written before changes)

# parameters whose values can be lost: every existing entry of a selected item, plus migration sources
backup_entries = []
for item in selected:
    backup_entries += item["delete"] + ([item["keep"]] if item.get("keep") else [])
by_cat = {}
for e in backup_entries:
    for cid in e["cat_ids"]:
        by_cat.setdefault(cid, []).append(e)

records = (
    []
)  # (name, guid, element id, is type, storage, value, type id of the instance)
instances_of_type = {}
for el in FilteredElementCollector(doc).WhereElementIsNotElementType():
    if el.Category is None:
        continue
    entries = by_cat.get(idv(el.Category.Id))
    if not entries:
        continue
    tid = el.GetTypeId()
    tid = idv(tid) if tid is not None and tid != ElementId.InvalidElementId else None
    if tid is not None:
        instances_of_type.setdefault(tid, []).append(idv(el.Id))
    for e in entries:
        v = read_value(el.get_Parameter(e["definition"]))
        if v:
            records.append(
                (
                    e["definition"].Name,
                    e["guid"],
                    idv(el.Id),
                    False,
                    v[0],
                    v[1],
                    tid,
                    idv(el.Category.Id),
                )
            )
for el in FilteredElementCollector(doc).WhereElementIsElementType():
    if el.Category is None:
        continue
    for e in by_cat.get(idv(el.Category.Id), []):
        v = read_value(el.get_Parameter(e["definition"]))
        if v:
            records.append(
                (
                    e["definition"].Name,
                    e["guid"],
                    idv(el.Id),
                    True,
                    v[0],
                    v[1],
                    None,
                    idv(el.Category.Id),
                )
            )

backup_path = save_log(
    "BMC_BackupParametri",
    {
        "model": doc.Title,
        "actions": [(i["name"], i["action"]) for i in selected],
        "fields": [
            "param",
            "guid",
            "element_id",
            "is_type",
            "storage",
            "value",
            "type_id",
            "category_id",
        ],
        "records": records,
    },
)
output.print_md(
    "Backup valori (%d) salvato prima delle modifiche: `%s`"
    % (len(records), backup_path)
)

# values on categories the PIR no longer binds cannot be written back: show them before confirming
restorable = dict(
    (i["name"], set(i["target_cat_ids"]))
    for i in selected
    if i["action"] in ("riassocia", "sostituisci", "elimina copie")
)
excluded = Counter()
for r in records:
    if r[0] in restorable and r[7] not in restorable[r[0]]:
        cat = doc.GetElement(ElementId(r[2])).Category
        excluded[(r[0], cat.Name if cat else "?")] += 1
if excluded:
    output.print_md(
        "## Valori su categorie non previste dal PIR (restano solo nel backup)"
    )
    output.print_table(
        table_data=[[n, c, k] for (n, c), k in sorted(excluded.items())],
        columns=["Parametro", "Categoria", "Valori"],
    )

if not forms.alert(
    "Applicare %d azioni sui parametri e ricompilare %d valori?\n"
    "%d valori su categorie non previste dal PIR resteranno solo nel backup.\n"
    "Backup: %s\nL'operazione si annulla con Ctrl+Z."
    % (len(selected), len(records), sum(excluded.values()), backup_path),
    yes=True,
    no=True,
):
    script.exit()

# ---------------------------------------------------------------- apply

log = {
    "model": doc.Title,
    "backup": backup_path,
    "done": [],
    "failed": [],
    "values": {},
}
stats = {}  # param -> Counter(restored, failed, conflicts, migrated)


def stat(name, key, n=1):
    stats.setdefault(name, Counter())[key] += n


def insert_from_txt(items):
    """Insert the given plan items from their TXT files (the shared parameter file is restored after)."""
    old_file = app.SharedParametersFilename
    try:
        for file_key, file_name in R.SHARED_FILES.items():
            todo = [i for i in items if i["file"] == file_key]
            if not todo:
                continue
            app.SharedParametersFilename = os.path.join(R.SHARED_DIR, file_name)
            dfile = app.OpenSharedParameterFile()
            definitions = {}
            for g in dfile.Groups:
                for d in g.Definitions:
                    definitions[d.Name] = d
            for i in todo:
                try:
                    ext = definitions[i["name"]]
                    group = i["delete"][0]["group"] if i["delete"] else GroupTypeId.Text
                    if not doc.ParameterBindings.Insert(
                        ext, make_binding(i["level"], i["bics"]), group
                    ):
                        raise Exception("Insert rifiutato da Revit")
                    log["done"].append([i["name"], i["action"]])
                except Exception as ex:
                    log["failed"].append([i["name"], i["action"], str(ex)])
    finally:
        app.SharedParametersFilename = old_file


def restore(targets_by_name):
    """Write the backup values back into the (new or rebound) parameters, only where they are empty.
    Values of the kept GUID are written first, the other copies only fill what is still empty.
    """
    ordered = sorted(
        records,
        key=lambda r: (
            0
            if (
                targets_by_name.get(r[0])
                and targets_by_name[r[0]]["keep"]
                and r[1] == targets_by_name[r[0]]["keep"]["guid"]
            )
            else 1
        ),
    )
    type_votes = {}
    for name, guid, eid, is_type, storage, value, tid, cid in ordered:
        item = targets_by_name.get(name)
        if item is None:
            continue
        try:
            if item["level"] == "T" and not is_type:
                if tid is not None:
                    type_votes.setdefault((name, tid), Counter())[(storage, value)] += 1
                continue
            ids = (
                [eid]
                if (item["level"] == "T") == is_type
                else instances_of_type.get(eid, [])
            )
            for target_id in ids:
                p = doc.GetElement(ElementId(target_id)).LookupParameter(name)
                if p is None:
                    stat(name, "categoria esclusa dal PIR (solo backup)")
                    continue
                if p.IsReadOnly or not is_empty(p):
                    continue
                write_value(p, storage, value)
                stat(name, "ripristinati")
        except Exception as ex:
            stat(name, "errori")
            log["failed"].append([name, "valore id %s" % eid, str(ex)])
    for (name, tid), votes in type_votes.items():
        (storage, value), n = votes.most_common(1)[0]
        if len(votes) > 1:
            stat(name, "conflitti")
        try:
            p = doc.GetElement(ElementId(tid)).LookupParameter(name)
            if p is not None and not p.IsReadOnly and is_empty(p):
                write_value(p, storage, value)
                stat(name, "ripristinati")
        except Exception as ex:
            stat(name, "errori")
            log["failed"].append([name, "valore tipo %s" % tid, str(ex)])


def migrate():
    for source, target, bics, conversion in R.MIGRATIONS:
        allowed = None
        if bics:
            allowed = set(
                idv(Category.GetCategory(doc, b).Id)
                for b in bics
                if Category.GetCategory(doc, b)
            )
        for name, guid, eid, is_type, storage, value, tid, cid in records:
            if name != source:
                continue
            el = doc.GetElement(ElementId(eid))
            if el is None or (
                allowed is not None
                and (el.Category is None or idv(el.Category.Id) not in allowed)
            ):
                continue
            p = el.LookupParameter(target)
            try:
                if p is not None and not p.IsReadOnly and is_empty(p):
                    write_value(p, storage, value, conversion)
                    stat(target, "migrati da %s" % source)
            except Exception as ex:
                stat(target, "errori")
                log["failed"].append([target, "migrazione id %s" % eid, str(ex)])


def reattach_fields():
    by_name = dict((i["name"], i) for i in selected)
    for sid, index, ftype, heading, hidden, name in sorted(
        field_moves, key=lambda m: (idv(m[0]), m[1])
    ):
        item = by_name[name]
        try:
            if item["keep"]:
                pid = item["keep"]["pid"]
            else:
                spe = SharedParameterElement.Lookup(
                    doc, Guid(R.SHARED_PARAMS[name]["guid"])
                )
                if spe is None:
                    raise Exception("parametro TXT non inserito")
                pid = spe.Id
            d = doc.GetElement(sid).Definition
            f = d.InsertField(ftype, pid, min(index, d.GetFieldCount()))
            f.ColumnHeading = heading
            f.IsHidden = hidden
            stat(name, "colonne abaco ripristinate")
        except Exception as ex:
            stat(name, "colonne abaco non ripristinate")
            log["failed"].append(
                [name, "colonna in %s" % doc.GetElement(sid).Name, str(ex)]
            )


t = Transaction(doc, "BMC - Riemissione parametri R06")
t.Start()
try:
    # 1. delete wrong-GUID copies, replaced and legacy parameters (values are in the backup)
    for item in selected:
        for e in item["delete"]:
            try:
                doc.Delete(e["pid"])
                log["done"].append([e["name"], "eliminato"])
            except Exception as ex:
                log["failed"].append([item["name"], "elimina", str(ex)])
    # 2. rebind the parameters that already have the TXT GUID
    for item in [i for i in selected if i["action"] == "riassocia"]:
        try:
            k = item["keep"]
            if not doc.ParameterBindings.ReInsert(
                k["definition"],
                make_binding(item["level"], item["bics"]),
                k["definition"].GetGroupTypeId(),
            ):
                raise Exception("ReInsert rifiutato da Revit")
            log["done"].append([item["name"], "riassocia"])
        except Exception as ex:
            log["failed"].append([item["name"], "riassocia", str(ex)])
    # 3. insert missing and replaced parameters from the TXT files
    insert_from_txt([i for i in selected if i["action"] in ("aggiungi", "sostituisci")])
    doc.Regenerate()
    # 4. values back, then migrations from the legacy parameters
    restore(
        dict(
            (i["name"], i)
            for i in selected
            if i["action"] in ("riassocia", "sostituisci", "elimina copie")
        )
    )
    if "elimina legacy" in chosen:
        migrate()
    # 5. schedule columns of the deleted copies point again to the final parameter
    reattach_fields()
except Exception as ex:
    t.RollBack()
    forms.alert(
        "Errore imprevisto, nessuna modifica applicata:\n%s" % ex, exitscript=True
    )
else:
    log["commit_status"] = commit(t)

log["values"] = dict((k, dict(v)) for k, v in stats.items())
log_path = save_log("BMC_RiemissioneParametri", log)

output.print_md("---")
if log["commit_status"] != "Committed":
    output.print_md(
        "# ATTENZIONE: transazione %s, il modello NON e' stato modificato"
        % log["commit_status"]
    )
output.print_md(
    "# Operazioni riuscite: %d - non riuscite: %d"
    % (len(log["done"]), len(log["failed"]))
)
if stats:
    output.print_table(
        table_data=[
            [n, ", ".join("%s %d" % kv for kv in sorted(c.items()))]
            for n, c in sorted(stats.items())
        ],
        columns=["Parametro", "Valori"],
    )
if log["failed"]:
    output.print_md("## Non riuscite")
    output.print_table(
        table_data=log["failed"][:MAX_ROWS],
        columns=["Parametro", "Operazione", "Errore"],
    )
output.print_md("Backup: `%s`  \nLog: `%s`" % (backup_path, log_path))
output.print_md("Poi: Model Data Export + Model Checker (00.4 e 06.x) per verificare.")
