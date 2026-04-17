"""Generate a professional BIM warning report for coordination meetings"""

__title__ = "Warning\nBIM Report"
__author__ = "Luca Rosati"

import clr
import System
import datetime
from collections import defaultdict

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import *

from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document

# --- Raccolta warning ---
doc_warn = doc.GetWarnings()

if not doc_warn or len(doc_warn) == 0:
    forms.alert("Nessun warning trovato nel modello. Ottimo lavoro!", exitscript=True)

# --- Raggruppamento per descrizione (tipo warning) ---
# Struttura: { desc: { severity, count, elements[] } }
warn_groups = {}

for w in doc_warn:
    desc = w.GetDescriptionText()
    sev = w.GetSeverity()
    sev_label = "ERRORE" if str(sev) == "Error" else "WARNING"

    if desc not in warn_groups:
        warn_groups[desc] = {"severity": sev_label, "count": 0, "elements": []}

    warn_groups[desc]["count"] += 1
    for eid in w.GetFailingElements():
        warn_groups[desc]["elements"].append(eid)

# --- Contatori globali ---
total_warn = len(doc_warn)
n_tipi_errore = sum(1 for g in warn_groups.values() if g["severity"] == "ERRORE")
n_tipi_warning = sum(1 for g in warn_groups.values() if g["severity"] == "WARNING")
n_tot_errore = sum(
    g["count"] for g in warn_groups.values() if g["severity"] == "ERRORE"
)
n_tot_warning = sum(
    g["count"] for g in warn_groups.values() if g["severity"] == "WARNING"
)

# --- Info progetto ---
proj_info = doc.ProjectInformation
proj_name = proj_info.Name if proj_info.Name else "N/D"
proj_number = proj_info.Number if proj_info.Number else "N/D"
now = datetime.datetime.now()
date_str = now.strftime("%d/%m/%Y")
time_str = now.strftime("%H:%M")

# --- Ordinamento per occorrenze (desc) ---
sorted_groups = sorted(warn_groups.items(), key=lambda x: x[1]["count"], reverse=True)

# --- Output ---
output = script.get_output()
output.set_title("Warning BIM Report")

output.print_md("# Warning BIM Report")
output.print_md(
    "**Progetto:** {}  |  **N. Progetto:** {}  |  **Data:** {}  |  **Warning Totali:** {}".format(
        proj_name, proj_number, date_str, total_warn
    )
)
output.print_md("---")

# Sommario severita
output.print_md("## Sommario")
output.print_md("| Severita | Tipi Distinti | Occorrenze Totali |")
output.print_md("|---|---|---|")
if n_tipi_errore > 0:
    output.print_md("| ERRORE | {} | {} |".format(n_tipi_errore, n_tot_errore))
output.print_md("| WARNING | {} | {} |".format(n_tipi_warning, n_tot_warning))
output.print_md("")

# Top 5 problemi prioritari
output.print_md("## Top 5 Problemi Prioritari")
output.print_md("| # | Tipo Warning | Severita | Occorrenze |")
output.print_md("|---|---|---|---|")
for i, (desc, data) in enumerate(sorted_groups[:5]):
    short_desc = desc[:80] + "..." if len(desc) > 80 else desc
    output.print_md(
        "| {} | {} | {} | {} |".format(
            i + 1, short_desc, data["severity"], data["count"]
        )
    )
output.print_md("")

# Dettaglio completo
output.print_md("## Dettaglio Completo per Tipo")
output.print_md("")

for desc, data in sorted_groups:
    sev = data["severity"]
    count = data["count"]
    elements = data["elements"]

    # Header tipo warning
    output.print_md("### {} — {} occorrenze".format(sev, count))
    output.print_md("**Descrizione:** {}".format(desc))

    # Elementi linkificati (max 5)
    if elements:
        output.print_md("**Elementi coinvolti:**")
        max_show = 5
        shown = elements[:max_show]
        remaining = len(elements) - max_show

        elem_parts = []
        for eid in shown:
            elem_parts.append(output.linkify(eid))

        elem_line = "  ".join(elem_parts)
        if remaining > 0:
            elem_line += "  *(e altri {} elementi)*".format(remaining)

        output.print_md(elem_line)
    output.print_md("")

# Footer
output.print_md("---")
output.print_md(
    "*Report generato il {} alle {} — VisionXtools Warning BIM Report*".format(
        date_str, time_str
    )
)
