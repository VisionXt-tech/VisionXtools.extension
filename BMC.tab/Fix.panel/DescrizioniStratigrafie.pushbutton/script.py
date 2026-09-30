"""Rebuild the Description of every layered type (walls, floors, ceilings, roofs) from its real layers"""

__title__ = "Descrizioni\nstratigrafie"
__author__ = "Luca Rosati"
__guida__ = "Ricompone la Description di tutti i tipi a strati del progetto (muri, pavimenti, controsoffitti, coperture): mantiene il testo prima di 'Composizione:' e riscrive la composizione dagli strati reali, con il nome completo del materiale (le descrizioni vecchie lo tagliavano con '...') e lo spessore in cm. Prima anteprima, poi applica; log JSON con i valori vecchi."

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# Flow: ask preview / apply -> compute (read only) -> preview -> [apply: one Transaction, Ctrl+Z undoes it,
# JSON log with old and new values]. Text rule: bmc_rules.stratigraphy_description (self-checked).
# Types with a layer without material are listed and skipped: their composition cannot be written.

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script

import bmc_rules as R
from bmc_utils import ask_preview_only, commit, end_preview, ename, save_log, to_mm

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

TYPE_CLASSES = [
    ("Muri", WallType),
    ("Pavimenti", FloorType),
    ("Controsoffitti", CeilingType),
    ("Coperture", RoofType),
]
MAX_ROWS = 300

PREVIEW_ONLY = ask_preview_only()

# ---------------------------------------------------------------- compute (read only)

changes, same, skipped = [], 0, []
for label, cls in TYPE_CLASSES:
    for element_type in FilteredElementCollector(doc).OfClass(cls):
        cs = element_type.GetCompoundStructure()
        if (
            cs is None or cs.LayerCount == 0
        ):  # curtain / stacked walls, in-place, no layers
            continue
        layers, missing = [], 0
        for layer in cs.GetLayers():
            material = doc.GetElement(layer.MaterialId)
            if material is None:
                missing += 1
                continue
            layers.append((ename(material), to_mm(layer.Width)))
        if missing:
            skipped.append(
                (label, ename(element_type), "%d strati senza materiale" % missing)
            )
            continue
        p = element_type.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION)
        if p is None or p.IsReadOnly:
            skipped.append((label, ename(element_type), "Description non scrivibile"))
            continue
        current = p.AsString() or ""
        new = R.stratigraphy_description(current, layers)
        if new == current:
            same += 1
        else:
            changes.append((label, element_type, current, new))

# ---------------------------------------------------------------- preview

output.print_md("# BMC - Descrizioni stratigrafie")
output.print_md(
    "**%d tipi da aggiornare**, %d gia' corretti, %d saltati."
    % (len(changes), same, len(skipped))
)
if changes:
    output.print_table(
        table_data=[
            [label, ename(t), current or "(vuota)", new]
            for label, t, current, new in changes[:MAX_ROWS]
        ],
        columns=["Categoria", "Tipo", "Description attuale", "Description nuova"],
    )
    if len(changes) > MAX_ROWS:
        output.print_md("... e altri %d (tutti nel log)." % (len(changes) - MAX_ROWS))
if skipped:
    output.print_md("## Tipi saltati")
    output.print_table(table_data=skipped, columns=["Categoria", "Tipo", "Motivo"])
end_preview(output, PREVIEW_ONLY)
if not changes:
    script.exit()
if not forms.alert(
    "Aggiornare la Description di %d tipi?" % len(changes), yes=True, no=True
):
    script.exit()

# ---------------------------------------------------------------- apply

log = {"model": doc.Title, "changes": [], "errors": []}
t = Transaction(doc, "BMC - Descrizioni stratigrafie")
t.Start()
try:
    for label, element_type, current, new in changes:
        try:
            element_type.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).Set(new)
            log["changes"].append(
                {
                    "category": label,
                    "type": ename(element_type),
                    "old": current,
                    "new": new,
                }
            )
        except Exception as ex:
            log["errors"].append({"type": ename(element_type), "error": str(ex)})
except Exception as ex:
    t.RollBack()
    forms.alert(
        "Errore imprevisto, nessuna modifica applicata:\n%s" % ex, exitscript=True
    )
log["commit_status"] = commit(t)
log_path = save_log("BMC_DescrizioniStratigrafie", log)

output.print_md("---")
if log["commit_status"] != "Committed":
    output.print_md(
        "# ATTENZIONE: transazione %s, il modello NON e' stato modificato"
        % log["commit_status"]
    )
else:
    output.print_md(
        "# Description aggiornate: %d - errori: %d"
        % (len(log["changes"]), len(log["errors"]))
    )
for e in log["errors"]:
    output.print_md("- `%s`: %s" % (e["type"], e["error"]))
output.print_md("Log (valori vecchi e nuovi): `%s`" % log_path)
