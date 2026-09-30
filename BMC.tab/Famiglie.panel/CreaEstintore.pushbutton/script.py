"""Build the portable fire extinguisher family (powder 6 kg, CO2 5 kg) inside the open Specialty Equipment template"""

__title__ = "Crea\nestintore"
__author__ = "Luca Rosati"
__guida__ = "Da lanciare nell'editor di famiglie, sul template Metric Specialty Equipment appena aperto: costruisce l'estintore portatile a parete (serbatoio, valvola, staffa, cartello UNI EN ISO 7010 F001 con Si/No) con 2 tipi, polvere 6 kg 34A 233B C e CO2 5 kg 113B. In pianta la geometria e' spenta e compare in rosso il simbolo del DM 30/11/1983 (quadrato con E, senza testo) in dimensione carta; se appare capovolto si spunta 'Simbolo ruotato 180'. Altezza dell'impugnatura come parametro di istanza, parametri informativi compilati, prova di tutti i tipi, salvataggio con il nome PGI."

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# PGI R06 3.11.5-3.11.6: family AR_ATS_ES1.00A_<descr>, types AR_ATS_ES1.0nA_<descr>_<diameter x height mm>.
# Category Specialty Equipment (decision 29/09/2026: PGI code ATS; Revit Fire Protection has no PGI code).
# Dimensions: Emme Italia datasheets, powder 6 kg cod. 21063-51 (D160 x 510, 9.4 kg), CO2 5 kg cod. 23052-1
# (D152 x 670, 12.0 kg), both UNI EN 3-7. Codice prevenzione incendi S.6.6.2.1: class A extinguishers on every
# floor, reach <= 30 m (Rvita A3/B1-B2), >= 21A; S.6.11: signs UNI EN ISO 7010.
# Plan symbol: DM 30/11/1983 Allegato B "Estintore portatile" = square with E, with the fire class and the
# extinguishing capacity written next to it (note of the table). Decision 29/09/2026: symbol inside the family,
# plan views only, red, 3D geometry hidden in plan/RCP; no text next to it (decision 29/09/2026).
#
# Level based, low parametric (user request): plan geometry fixed, heights driven by parameters. Origin = wall
# face (template front/back centre plane), extinguisher towards -Y (front view). A cylinder has no planar face
# to lock, so each type has its own tank + valve, shown by the type Yes/No "Tipo <short name>".
# The symbol is a nested Generic Annotation (paper size at every scale), the same for every type.
# Readability: a nested annotation turns with the instance and a family cannot read its own rotation, so there is
# also a variant drawn turned 180 degrees about the square centre; the instance Yes/No "Simbolo ruotato 180"
# swaps them (host 180 + variant 180 = upright). E is symmetric about its horizontal axis, so the same switch
# also fixes an instance mirrored about a vertical axis.
# ponytail: one body per type; fine for 2-3 types, a nested family with a radius beyond that.
# The switch is set by hand; a command reading FacingOrientation could set it on every instance if needed.

import os
import tempfile

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script

import bmc_rules as R
from bmc_family import (
    circle,
    extrude,
    ft,
    length_param,
    level_plane,
    plan_view,
    rectangle,
    to_mm,
)

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
output = script.get_output()

BIC = BuiltInCategory.OST_SpecialityEquipment
CAT = R.CAT_CODE[BIC]  # ATS
TEMPLATE = "Metric Specialty Equipment.rft"
SYMBOL_TEMPLATE = "Metric Generic Annotation.rft"

# (short name, agent, charge, fire class, diameter, height, weight kg, manufacturer code, cylinder), mm
TYPES = [
    (
        "polvere 6kg",
        "a polvere ABC",
        "6 kg",
        "34A 233B C",
        160,
        510,
        "9,4",
        "21063-51",
        "serbatoio in acciaio",
    ),
    (
        "CO2 5kg",
        "ad anidride carbonica CO2",
        "5 kg",
        "113B",
        152,
        670,
        "12,0",
        "23052-1",
        "bombola in lega di alluminio",
    ),
]
# mm; not in the datasheets: declared defaults, type parameters
HEAD_D, HEAD_H = (
    50,
    110,
)  # valve head diameter (fixed) and height above the tank shoulder
BRACKET_W, BRACKET_T, BRACKET_H = (
    80,
    20,
    150,
)  # wall bracket: width, depth from the wall, height
SIGN_W, SIGN_T, SIGN_H = 200, 2, 200  # F001 sign
HANDLE = 1100  # handle height from the level (common practice, S.6 sets no height)
SIGN_Z = 1800  # sign bottom from the level
# plan symbol, paper mm: square side, E letter half width and half height
SYM, E_W, E_H = 5.0, 1.2, 1.5
SYMBOL_SUBCAT = "Simbolo estintore"
RED = Color(255, 0, 0)

FAMILY_NAME = "AR_%s_ES1.00A_Estintore portatile" % CAT
P_HANDLE, P_H, P_HEAD, P_BRACKET = (
    "Altezza impugnatura",
    "Altezza estintore",
    "Altezza valvola",
    "Altezza staffa",
)
Z_BASE, Z_SHOULDER, Z_BRACKET = "Quota base", "Quota spalla", "Quota staffa"
P_SIGN_Z, P_SIGN_H, Z_SIGN_TOP = (
    "Quota cartello",
    "Altezza cartello",
    "Quota testa cartello",
)
P_SHOW_SIGN, P_MATERIAL = "Mostra cartello", "Materiale estintore"
P_TURNED, P_UPRIGHT = "Simbolo ruotato 180", "Simbolo dritto"
VARIANTS = (False, True)  # symbol drawn upright, symbol drawn turned 180 degrees


def mark(i):
    return "ES1.%02dA" % (i + 1)


def type_name(i):
    short, _a, _c, _k, d, h = TYPES[i][:6]
    return "AR_%s_%s_Estintore %s_%dx%d" % (CAT, mark(i), short, d, h)


def type_param(i):
    return "Tipo %s" % TYPES[i][0]


def symbol_name(turned):
    # nested, not shared: never listed in the project
    return "Simbolo estintore%s" % (" ruotato" if turned else "")


def type_values(i):
    _s, agent, charge, fire, d, h, kg, code, cylinder = TYPES[i]
    return [
        (BuiltInParameter.ALL_MODEL_TYPE_MARK, mark(i)),
        (BuiltInParameter.KEYNOTE_PARAM, "AR_%s_%s" % (CAT, mark(i))),
        (
            BuiltInParameter.ALL_MODEL_TYPE_COMMENTS,
            "ESTINTORE %s" % TYPES[i][0].upper(),
        ),
        (
            BuiltInParameter.ALL_MODEL_DESCRIPTION,
            "Estintore portatile %s tipo Emme Italia cod. %s o equivalente, carica nominale %s, classe di "
            "fuoco %s, conforme UNI EN 3-7, %s rosso RAL 3000, peso totale %s kg, diametro %d mm, altezza "
            "%d mm, installato a parete su staffa con cartello UNI EN ISO 7010 F001."
            % (agent, code, charge, fire, cylinder, kg, d, h),
        ),
        (BuiltInParameter.ALL_MODEL_MANUFACTURER, "Emme Italia"),
        (BuiltInParameter.ALL_MODEL_MODEL, code),
    ]


# names checked with the same rules as the Model Checker (05.2)
assert R.RE_FAMILY.match(FAMILY_NAME)
for _i in range(len(TYPES)):
    assert R.RE_TYPE.match(type_name(_i)), type_name(_i)


def template_path(name, sub):
    folders = [
        os.path.join(app.FamilyTemplatePath or "", sub),
        os.path.join(
            os.environ.get("ProgramData", r"C:\ProgramData"),
            "Autodesk",
            "RVT %s" % app.VersionNumber,
            "Family Templates",
            "English",
            sub,
        ),
    ]
    for folder in folders:
        path = os.path.join(folder, name)
        if os.path.exists(path):
            return path
    path = forms.pick_file(file_ext="rft", title="Scegli il template '%s'" % name)
    if not path:
        script.exit()
    return path


def moved(loop, dy):
    """Copy of a closed loop moved along Y (mm)."""
    t = Transform.CreateTranslation(XYZ(0, ft(dy), 0))
    out = CurveArray()
    for c in loop:
        out.Append(c.CreateTransformed(t))
    return out


def z_range_mm(element):
    bb = element.get_BoundingBox(None)
    return to_mm(bb.Min.Z), to_mm(bb.Max.Z)


def hide_in_plan(element):
    """3D geometry shown in elevations, sections and 3D, never in plan / RCP (cut or not)."""
    vis = FamilyElementVisibility(FamilyElementVisibilityType.Model)
    vis.IsShownInTopBottom = False
    vis.IsShownInPlanRCPCut = False
    element.SetVisibility(vis)


# ---------------------------------------------------------------- nested plan symbols


def symbol_segments():
    """Symbol in paper mm, upright: square back edge on the origin (wall face), towards -Y, E inside."""
    h, cy = SYM / 2.0, -SYM / 2.0
    return [
        (-h, 0, h, 0),
        (h, 0, h, -SYM),
        (h, -SYM, -h, -SYM),
        (-h, -SYM, -h, 0),
        (-E_W, cy + E_H, -E_W, cy - E_H),  # E: stem
        (-E_W, cy + E_H, E_W, cy + E_H),  # E: top bar
        (-E_W, cy, E_W * 0.7, cy),  # E: middle bar
        (-E_W, cy - E_H, E_W, cy - E_H),  # E: bottom bar
    ]


def turn(x, y):
    """180 degrees about the square centre (0, -SYM/2): the square maps onto itself."""
    return -x, -SYM - y


def draw_symbol(ndoc, turned):
    """DM 30/11/1983 portable extinguisher, red lines only; turned = drawn 180 degrees about the square centre.
    Runs inside a transaction of the annotation document."""
    view = [
        v
        for v in FilteredElementCollector(ndoc).OfClass(View)
        if not v.IsTemplate
        and v.ViewType
        not in (ViewType.ProjectBrowser, ViewType.SystemBrowser, ViewType.Internal)
    ][0]
    # the template may carry an instruction note: the symbol must show only what we draw
    for note in list(FilteredElementCollector(ndoc).OfClass(TextNote)):
        ndoc.Delete(note.Id)
    sub = ndoc.Settings.Categories.NewSubcategory(
        ndoc.OwnerFamily.FamilyCategory, SYMBOL_SUBCAT
    )
    sub.LineColor = RED
    line_style = sub.GetGraphicsStyle(GraphicsStyleType.Projection)
    for x1, y1, x2, y2 in symbol_segments():
        if turned:
            (x1, y1), (x2, y2) = turn(x1, y1), turn(x2, y2)
        curve = ndoc.FamilyCreate.NewDetailCurve(
            view, Line.CreateBound(XYZ(ft(x1), ft(y1), 0), XYZ(ft(x2), ft(y2), 0))
        )
        curve.LineStyle = line_style


def symbol_family(fdoc, turned, template):
    """Nested annotation family, loaded into fdoc (reused if a previous run already loaded it)."""
    name = symbol_name(turned)
    for fam in FilteredElementCollector(fdoc).OfClass(Family):
        if fam.Name == name:
            return fam
    ndoc = app.NewFamilyDocument(template)
    try:
        t = Transaction(ndoc, "BMC - Simbolo estintore")
        t.Start()
        try:
            draw_symbol(ndoc, turned)
        except Exception:
            t.RollBack()
            raise
        if str(t.Commit()) != "Committed":
            raise Exception("transazione del simbolo non confermata")
        path = os.path.join(
            tempfile.gettempdir(), name + ".rfa"
        )  # the file name is the family name
        opts = SaveAsOptions()
        opts.OverwriteExistingFile = True
        ndoc.SaveAs(path, opts)
        return ndoc.LoadFamily(fdoc)
    finally:
        ndoc.Close(False)


# ---------------------------------------------------------------- host family


def visible(e):
    return e.get_Parameter(BuiltInParameter.IS_VISIBLE_PARAM).AsInteger() == 1


def flex_test(fdoc, bodies, symbols, p_turned):
    """Every type made current in turn, symbol upright and turned: only its own tank, valve and the right
    symbol visible; diameter, base and top at the handle. Reads the elements' Visible values, so it checks
    the associations and formulas too. Rolled back."""
    fm = fdoc.FamilyManager
    types = dict((t.Name, t) for t in fm.Types)
    st = SubTransaction(fdoc)
    st.Start()
    try:
        results = []
        for i in range(len(TYPES)):
            fm.CurrentType = types[type_name(i)]
            for turned in VARIANTS:
                fm.Set(p_turned, 1 if turned else 0)
                fdoc.Regenerate()
                checks = []
                for j in range(len(TYPES)):
                    checks += [
                        ("serbatoio %s" % mark(j), bodies[j][0], i == j),
                        ("valvola %s" % mark(j), bodies[j][1], i == j),
                    ]
                checks += [(symbol_name(v), symbols[v], v == turned) for v in VARIANTS]
                for what, e, expected in checks:
                    shown = visible(e)
                    results.append(
                        (
                            "%s%s: %s visibile"
                            % (mark(i), " ruotato" if turned else "", what),
                            shown,
                            expected,
                            shown == expected,
                        )
                    )
            d, h = TYPES[i][4], TYPES[i][5]
            tank, head = bodies[i]
            bb = tank.get_BoundingBox(None)
            got = (
                to_mm(bb.Max.X - bb.Min.X),
                z_range_mm(tank)[0],
                z_range_mm(head)[1],
            )
            expected = (d, HANDLE - h, HANDLE)
            ok = all(abs(g - e) < 1.0 for g, e in zip(got, expected))
            results.append(
                ("%s diametro / base / impugnatura" % mark(i), got, expected, ok)
            )
    except Exception as ex:
        results = [("Flessione", str(ex), "-", False)]
    finally:
        st.RollBack()
    return results


def build(fdoc, symbol_families):
    """Everything inside the family (open transaction). Returns (flex results, notes)."""
    fm = fdoc.FamilyManager
    notes, step = [], ""
    try:
        step = "tipo e parametri"
        if fm.Types.Size == 0:
            fm.NewType(type_name(0))
        else:
            fm.RenameCurrentType(type_name(0))
        p_h = length_param(fm, P_H, TYPES[0][5])
        length_param(fm, P_HEAD, HEAD_H)
        length_param(fm, P_BRACKET, BRACKET_H)
        length_param(fm, P_SIGN_H, SIGN_H)
        length_param(fm, P_HANDLE, HANDLE, instance=True)
        length_param(fm, P_SIGN_Z, SIGN_Z, instance=True)
        z_base = length_param(
            fm, Z_BASE, formula="%s - %s" % (P_HANDLE, P_H), instance=True
        )
        z_shoulder = length_param(
            fm, Z_SHOULDER, formula="%s - %s" % (P_HANDLE, P_HEAD), instance=True
        )
        z_bracket = length_param(
            fm, Z_BRACKET, formula="%s - %s" % (Z_SHOULDER, P_BRACKET), instance=True
        )
        z_sign_top = length_param(
            fm, Z_SIGN_TOP, formula="%s + %s" % (P_SIGN_Z, P_SIGN_H), instance=True
        )
        p_handle, p_sign_z = fm.get_Parameter(P_HANDLE), fm.get_Parameter(P_SIGN_Z)
        p_show = fm.AddParameter(
            P_SHOW_SIGN, GroupTypeId.Graphics, SpecTypeId.Boolean.YesNo, True
        )
        fm.Set(p_show, 1)
        p_turned = fm.AddParameter(
            P_TURNED, GroupTypeId.Graphics, SpecTypeId.Boolean.YesNo, True
        )
        fm.Set(p_turned, 0)
        p_mat = fm.AddParameter(
            P_MATERIAL, GroupTypeId.Materials, SpecTypeId.Reference.Material, False
        )
        p_vis = [
            fm.AddParameter(
                type_param(i), GroupTypeId.Graphics, SpecTypeId.Boolean.YesNo, False
            )
            for i in range(len(TYPES))
        ]
        p_upright = fm.AddParameter(
            P_UPRIGHT, GroupTypeId.Graphics, SpecTypeId.Boolean.YesNo, True
        )
        fm.SetFormula(p_upright, "not(%s)" % P_TURNED)
        fdoc.Regenerate()

        step = "estrusioni"
        plane = level_plane(fdoc)
        solids = []
        bodies = []
        for i, p in enumerate(p_vis):
            axis = -(BRACKET_T + TYPES[i][4] / 2.0)
            tank = extrude(
                fdoc, plane, [moved(circle(TYPES[i][4]), axis)], z_base, z_shoulder
            )
            head = extrude(
                fdoc, plane, [moved(circle(HEAD_D), axis)], z_shoulder, p_handle
            )
            for e in (tank, head):
                fm.AssociateElementParameterToFamilyParameter(
                    e.get_Parameter(BuiltInParameter.IS_VISIBLE_PARAM), p
                )
                fm.AssociateElementParameterToFamilyParameter(
                    e.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM), p_mat
                )
            bodies.append((tank, head))
            solids += [tank, head]
        solids.append(
            extrude(
                fdoc,
                plane,
                [moved(rectangle(BRACKET_W, BRACKET_T), -BRACKET_T / 2.0)],
                z_bracket,
                z_shoulder,
            )
        )
        sign = extrude(
            fdoc,
            plane,
            [moved(rectangle(SIGN_W, SIGN_T), -SIGN_T / 2.0)],
            p_sign_z,
            z_sign_top,
        )
        fm.AssociateElementParameterToFamilyParameter(
            sign.get_Parameter(BuiltInParameter.IS_VISIBLE_PARAM), p_show
        )
        solids.append(sign)

        step = "geometria spenta in pianta"
        for e in solids:
            hide_in_plan(e)

        step = "simboli in pianta"
        view = plan_view(fdoc)
        symbols = {}
        for turned in VARIANTS:
            symbol = fdoc.GetElement(
                list(symbol_families[turned].GetFamilySymbolIds())[0]
            )
            if not symbol.IsActive:
                symbol.Activate()
            inst = fdoc.FamilyCreate.NewFamilyInstance(XYZ.Zero, symbol, view)
            fm.AssociateElementParameterToFamilyParameter(
                inst.get_Parameter(BuiltInParameter.IS_VISIBLE_PARAM),
                p_turned if turned else p_upright,
            )
            symbols[turned] = inst
        fdoc.Regenerate()

        step = "tipi"
        first = fm.CurrentType
        for i, t in enumerate(TYPES):
            if i > 0:
                fm.NewType(type_name(i))  # copies the current values, becomes current
            fm.Set(p_h, ft(t[5]))
            for j, p in enumerate(p_vis):
                fm.Set(p, 1 if i == j else 0)
            for bip, value in type_values(i):
                fp = fm.get_Parameter(bip)
                if fp is None:
                    if i == 0:
                        notes.append(
                            "Parametro di tipo non presente nel template: %s" % bip
                        )
                else:
                    fm.Set(fp, value)
        fm.CurrentType = first
        fdoc.Regenerate()

        step = "prova dei tipi"
        flex = flex_test(fdoc, bodies, symbols, p_turned)
    except Exception as ex:
        raise Exception("%s: %s" % (step, ex))
    return flex, notes


# ---------------------------------------------------------------- checks (read only)

if not doc.IsFamilyDocument:
    forms.alert(
        "Apri prima il template '%s' (File > Nuovo > Famiglia) e lancia il comando dall'editor di famiglie."
        % TEMPLATE,
        exitscript=True,
    )
category = doc.OwnerFamily.FamilyCategory
if category is None or category.Id != Category.GetCategory(doc, BIC).Id:
    forms.alert(
        "Categoria della famiglia: '%s'. Serve 'Attrezzature speciali' (template '%s'): nessuna modifica."
        % (category.Name if category else "nessuna", TEMPLATE),
        exitscript=True,
    )
if doc.OwnerFamily.FamilyPlacementType != FamilyPlacementType.OneLevelBased:
    forms.alert(
        "La famiglia aperta non e' basata su livello: serve il template '%s'. Nessuna modifica."
        % TEMPLATE,
        exitscript=True,
    )
if doc.FamilyManager.get_Parameter(P_H):
    forms.alert(
        "La famiglia contiene gia' un estintore: il comando lavora solo su un template appena aperto "
        "(File > Nuovo > Famiglia > %s). Nessuna modifica." % TEMPLATE,
        exitscript=True,
    )
if not forms.alert(
    "Categoria verificata: %s, basata su livello.\nCreare la famiglia %s con i tipi\n%s?"
    % (category.Name, FAMILY_NAME, "\n".join(type_name(i) for i in range(len(TYPES)))),
    yes=True,
    no=True,
):
    script.exit()

# ---------------------------------------------------------------- build

# nested symbols first: LoadFamily cannot run while a transaction is open on the host
symbol_template = template_path(SYMBOL_TEMPLATE, "Annotations")
try:
    symbol_families = dict(
        (turned, symbol_family(doc, turned, symbol_template)) for turned in VARIANTS
    )
except Exception as ex:
    forms.alert(
        "Simboli in pianta non creati, famiglia non modificata:\n%s" % ex,
        exitscript=True,
    )

t = Transaction(doc, "BMC - Crea estintore")
t.Start()
try:
    flex, notes = build(doc, symbol_families)
except Exception as ex:
    t.RollBack()
    forms.alert(
        "Errore in '%s'\nFamiglia non modificata (restano caricati solo i simboli nidificati)."
        % ex,
        exitscript=True,
    )
status = str(t.Commit())
if status != "Committed":
    forms.alert("Transazione %s: famiglia non modificata." % status, exitscript=True)

# ---------------------------------------------------------------- save (optional) and report

rfa = forms.save_file(
    file_ext="rfa", default_name=FAMILY_NAME, title="Salva la famiglia con il nome PGI"
)
if rfa:
    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True  # the save dialog already asked about overwriting
    doc.SaveAs(rfa, opts)

output.print_md("# BMC - Crea estintore")
output.print_table(
    table_data=[
        [
            "Famiglia",
            (
                FAMILY_NAME
                if rfa
                else "%s (non salvata: il nome famiglia e' il nome del file)"
                % FAMILY_NAME
            ),
        ],
        ["Categoria", "Attrezzature speciali (%s), basata su livello" % CAT],
        ["File", rfa or "-"],
        [
            "Origine",
            "filo muro (piano centrale fronte/retro), estintore verso il fronte",
        ],
        [
            "Altezze (istanza)",
            "'%s' %d mm, '%s' %d mm" % (P_HANDLE, HANDLE, P_SIGN_Z, SIGN_Z),
        ],
        ["Cartello F001", "%dx%d mm, '%s' (istanza)" % (SIGN_W, SIGN_H, P_SHOW_SIGN)],
        [
            "Pianta",
            "geometria spenta; simbolo DM 30/11/1983 rosso (quadrato %g mm carta con E), "
            "sottocategoria '%s' delle Annotazioni generiche, '%s' (istanza) per leggerlo dritto"
            % (SYM, SYMBOL_SUBCAT, P_TURNED),
        ],
    ],
    columns=["Voce", "Valore"],
)
output.print_table(
    table_data=[
        [type_name(i), t[1], t[2], t[3], "%d x %d" % (t[4], t[5]), t[6], t[7]]
        for i, t in enumerate(TYPES)
    ],
    columns=[
        "Tipo",
        "Agente",
        "Carica",
        "Classe",
        "D x H (mm)",
        "Peso (kg)",
        "Cod. Emme",
    ],
)
output.print_md(
    "## Prova dei tipi (ogni tipo reso corrente, simbolo dritto e ruotato, poi annullata)"
)
bad = [f for f in flex if not f[3]]
output.print_md(
    "**%d controlli su %d OK.**" % (len(flex), len(flex))
    if not bad
    else "**ATTENZIONE: %d controlli su %d NON OK:**" % (len(bad), len(flex))
)
if bad:
    output.print_table(
        table_data=[[name, str(got), str(exp)] for name, got, exp, _ok in bad],
        columns=["Controllo", "Ottenuto", "Atteso"],
    )
for note in notes:
    output.print_md("- %s" % note)
output.print_md(
    "**Valori assunti (parametri di tipo, da allineare al prodotto scelto):** '%s' %d mm, diametro valvola %d mm, "
    "staffa %dx%dx%d mm. Materiale: assegnalo in '%s'.\n\n"
    "**Simbolo in pianta:** si legge dritto con l'estintore su un muro orizzontale con il fronte verso il "
    "basso della vista, e ruotato di 90 gradi in senso antiorario (lettura da destra). Se appare capovolto "
    "o con la E specchiata, spunta '%s' sull'istanza. Ruota gli estintori, non specchiarli.\n\n"
    "**Posizionamento (Codice prevenzione incendi S.6):** almeno un estintore di classe A per piano o "
    "compartimento, distanza di raggiungimento <= 30 m (Rvita A3, B1, B2), preferibilmente lungo le vie di "
    "esodo e vicino alle uscite; CO2 vicino ai quadri elettrici. Inserisci con il punto sul filo muro e "
    "ruota con la barra spaziatrice."
    % (
        P_HEAD,
        HEAD_H,
        HEAD_D,
        BRACKET_W,
        BRACKET_T,
        BRACKET_H,
        P_MATERIAL,
        P_TURNED,
    )
)
