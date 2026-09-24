"""List every button of the BMC tab with what it does, read live from the tab folders"""

__title__ = "Guida\nBMC"
__author__ = "Luca Rosati"
__guida__ = "Questa guida: elenca i pulsanti della tab BMC con una sintesi di cosa fanno. Si aggiorna da sola."

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source. Read only.
# Nothing is hard-coded: panels and buttons are read from the BMC.tab folders in their bundle.yaml order,
# the text of each button is its __guida__ line (fallback: the script docstring), parsed without running it.
# A new button appears here as soon as its folder exists: give it a __guida__ line in Italian.

import base64
import os

import _ast

from pyrevit import script

output = script.get_output()

TAB = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
PANEL_NOTE = {
    "Info": ("#62408c", "informazioni"),
    "Check": ("#1f5ea8", "sola lettura: non modifica il modello"),
    "Parametri": (
        "#1b8060",
        "scrive valori nei parametri: anteprima, conferma, Ctrl+Z, log in Documenti",
    ),
    "Fix": (
        "#d26816",
        "modifica nomi e struttura del modello: anteprima, conferma, Ctrl+Z, log in Documenti",
    ),
}
BUNDLE_KINDS = (".pushbutton", ".pulldown", ".splitbutton", ".stack")


def layout(folder):
    """Names listed under 'layout:' in bundle.yaml (minimal reader, no yaml dependency)."""
    path = os.path.join(folder, "bundle.yaml")
    names, inside = [], False
    if os.path.exists(path):
        for line in open(path):
            if line.strip().startswith("layout:"):
                inside = True
            elif inside and line.strip().startswith("- "):
                names.append(line.strip()[2:].strip())
            elif inside and line.strip():
                inside = False
    return names


def ordered(folder, suffixes):
    items = [
        n
        for n in os.listdir(folder)
        if n.endswith(suffixes) and os.path.isdir(os.path.join(folder, n))
    ]
    order = layout(folder)
    key = lambda n: (
        (
            order.index(n.rsplit(".", 1)[0])
            if n.rsplit(".", 1)[0] in order
            else len(order)
        ),
        n,
    )
    return sorted(items, key=key)


def describe(button_dir):
    """(title, text) from script.py without executing it."""
    path = os.path.join(button_dir, "script.py")
    title = os.path.basename(button_dir).rsplit(".", 1)[0]
    if not os.path.exists(path):
        return title, "(nessuno script)"
    tree = compile(open(path).read(), path, "exec", _ast.PyCF_ONLY_AST)
    values, doc = {}, ""
    for i, node in enumerate(tree.body):
        if i == 0 and isinstance(node, _ast.Expr) and isinstance(node.value, _ast.Str):
            doc = node.value.s
        if isinstance(node, _ast.Assign) and isinstance(node.value, _ast.Str):
            for target in node.targets:
                if isinstance(target, _ast.Name):
                    values[target.id] = node.value.s
    return (
        values.get("__title__", title).replace("\n", " "),
        values.get("__guida__") or doc,
    )


def icon_html(button_dir):
    path = os.path.join(button_dir, "icon.png")
    if not os.path.exists(path):
        return ""
    data = base64.b64encode(open(path, "rb").read())
    return '<img src="data:image/png;base64,%s" width="40" height="40"/>' % data


def cell(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


html = ['<h1 style="margin-bottom:4px">BMC - guida dei pulsanti</h1>']
html.append(
    "<p>Il <b>colore</b> dell'icona dice quanto il pulsante tocca il modello, il <b>simbolo</b> cosa fa. "
    "Regole di progetto: PGI R06 in <code>BMC.tab/lib/bmc_rules.py</code>.</p>"
)
for panel in ordered(TAB, (".panel",)):
    name = panel.rsplit(".", 1)[0]
    colour, note = PANEL_NOTE.get(name, ("#2b3a55", ""))
    html.append(
        '<h2 style="border-left:8px solid %s;padding-left:8px;margin-top:18px">%s'
        '<span style="font-size:13px;font-weight:normal"> - %s</span></h2>'
        % (colour, cell(name), cell(note))
    )
    html.append('<table style="border-collapse:collapse;width:100%">')
    for button in ordered(os.path.join(TAB, panel), BUNDLE_KINDS):
        button_dir = os.path.join(TAB, panel, button)
        title, text = describe(button_dir)
        html.append(
            '<tr><td style="width:52px;padding:6px 0;vertical-align:top">%s</td>'
            '<td style="padding:6px 8px;vertical-align:top;width:200px"><b>%s</b></td>'
            '<td style="padding:6px 8px;vertical-align:top">%s</td></tr>'
            % (icon_html(button_dir), cell(title), cell(text))
        )
    html.append("</table>")

output.set_title("BMC - Guida")
output.print_html("".join(html))
