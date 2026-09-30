"""Build the parametric roof access hatch (Gorter RHT, 4 sizes) inside the open face based family template"""

__title__ = "Crea botola\ncopertura"
__author__ = "Luca Rosati"
__guida__ = "Da lanciare nell'editor di famiglie, sul template Metric Generic Model face based appena aperto: imposta la categoria Attrezzature speciali e costruisce la botola di accesso alla copertura tipo Gorter RHT parametrica (controtelaio, telaio, converse con risvolto, guarnizione, coperchio, cassonetto della scala retrattile, area di manovra in pianta, foro nell'ospite) con 4 tipi conformi al Decreto Regione Lombardia 119/2009, prova la flessione su tutti i tipi, salva con il nome PGI e avvia Taglia geometria."

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# PGI R06 3.11.5-3.11.6: family AR_ATS_BT1.00A_<descr>, types AR_ATS_BT1.0nA_<descr>_<clear width x length mm>,
# progressive by size. Category Specialty Equipment (decision 28/09/2026: technical opaque hatch, not a window).
# Product: Gorter RHT standard (brochure IT 2026): clear openings 700x1000, 700x1400, 1000x1000, 1000x1500,
# Uw <= 0.276 W/m2K, max roof pitch 30 deg (BMC roof: 18.02 deg). Decreto Regione Lombardia 119/2009 point 1:
# inclined opening >= 0.50 m2 with shorter side >= 0.70 m. Positioning criteria and sources:
# docs/20260928/Botola_normativa_posizionamento.md.
#
# Face based template: the family sits on the host face (roof now, floor later), local Z = face normal, the host
# spans Z < 0. Parts, all locked to labelled reference planes:
#   controtelaio  clear opening -> outer frame, through the host (-HOST_THK .. 0)
#   telaio        clear opening -> outer frame, curb above the face (0 .. Altezza telaio)
#   converse      flashing plate outer frame -> converse (0 .. Spessore converse) + upstand on the curb
#   guarnizione   gasket on the curb top, under the lid
#   coperchio     lid overhanging the frame
#   area manovra  1 mm plate outside the flashing, dashed subcategory, plan only, Yes/No instance parameter
#   foro          void outer frame x -HOLE_DEPTH..+HOLE_MARGIN: cuts the host after Cut Geometry (manual click,
#                 SolidSolidCutUtils refuses face based families)
# Model lines cannot be joined through the API (a rectangle of lines would not flex): the manoeuvre area is a
# thin plate. Symbolic lines would not show in plan on a sloped host (view not parallel to the face).
# ponytail: HOST_THK fixed at the BMC roof (329 mm); make it an instance parameter if hosts of other
# thicknesses appear. Tags: no API creates labels in a tag family, the tag is made by hand (see report).

import clr

clr.AddReference("System")
from System.Collections.Generic import List

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import PostableCommand, RevitCommandId

from pyrevit import forms
from pyrevit import script

import bmc_rules as R
from bmc_family import (
    Planes,
    extrude,
    ft,
    length_param,
    level_plane,
    plan_view,
    rectangle,
    same,
    size_mm,
)

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

BIC = BuiltInCategory.OST_SpecialityEquipment
CAT = R.CAT_CODE[BIC]  # ATS
TEMPLATE = "Metric Generic Model face based.rft"

# (Gorter model, clear width, clear length) mm, by size: type marks BT1.01A..BT1.04A
TYPES = [
    ("RHT7010", 700, 1000),
    ("RHT7014", 700, 1400),
    ("RHT1010", 1000, 1000),
    ("RHT1015", 1000, 1500),
]
# Gorter scissor stair per type (decision 28/09/2026: closed stair casing only, ceiling < 3000 mm): (stair,
# casing width, casing length, predefined Gorter combination). RHT1010 has no predefined combination: prepared
# with the Small+ casing (700x1000 fits the 1000x1000 opening), to be confirmed by Gorter.
STAIRS = [
    ("Small+", 700, 1000, True),
    ("Large", 700, 1200, True),
    ("Small+", 700, 1000, False),
    ("XL", 1000, 1300, True),
]
CASE_H = 250  # stair casing depth under the host (Gorter structural opening 700x1200x250, 1000x1300x250)
# mm; not in the public Gorter data: declared defaults, type parameters to align to the Gorter drawing
PROFILE = 100  # frame profile width
CURB = 300  # curb height above the host face
GASKET = 10  # gasket thickness
LID = 80  # lid thickness
LID_OVER = 30  # lid overhang beyond the frame
FLASH_T = 5  # flashing sheet thickness (plate and upstand)
FLASH_W = 150  # flashing plate width beyond the frame
FLASH_H = 150  # flashing upstand height on the curb
AREA = 600  # free manoeuvre area beyond the frame (docs/20260928/Botola_normativa_posizionamento.md)
HOST_THK = 329  # BMC roof CS1.01A: sub-frame depth (see ponytail note)
HOLE_DEPTH, HOLE_MARGIN = 500, 10  # void below / above the host face
MIN_AREA, MIN_SIDE = 0.50, 700  # Decreto Regione Lombardia 119/2009, inclined opening

FAMILY_NAME = "AR_%s_BT1.00A_Botola accesso copertura" % CAT
SUB_AREA = "Botola - area di manovra"
P_W, P_L, P_PROF = "Larghezza luce", "Lunghezza luce", "Profilo telaio"
P_W_OUT, P_L_OUT, P_SURF = (
    "Larghezza esterna",
    "Lunghezza esterna",
    "Superficie luce netta",
)
P_FLASH_T, P_FLASH_W, P_FLASH_H = (
    "Spessore converse",
    "Larghezza converse",
    "Altezza risvolto converse",
)
P_OVER, P_AREA = "Sporgenza coperchio", "Area di manovra"
P_CURB, P_GASKET, P_LID = "Altezza telaio", "Spessore guarnizione", "Spessore coperchio"
Z_LID_BASE, Z_LID = "Quota base coperchio", "Quota coperchio"
P_SHOW_AREA = "Mostra area di manovra"
P_CASE_W, P_CASE_L = "Larghezza cassonetto scala", "Lunghezza cassonetto scala"
MATERIALS = [  # (parameter, parts)
    ("Materiale telaio", ["Telaio", "Coperchio", "Cassonetto scala"]),
    ("Materiale controtelaio", ["Controtelaio"]),
    ("Materiale converse", ["Converse", "Risvolto converse"]),
    ("Materiale guarnizione", ["Guarnizione"]),
]


def mark(i):
    return "BT1.%02dA" % (i + 1)


def type_name(i):
    _model, w, l = TYPES[i]
    return "AR_%s_%s_Botola accesso copertura_%dx%d" % (CAT, mark(i), w, l)


def stair_text(i):
    stair, cw, cl, predefined = STAIRS[i]
    if predefined:
        return "con scala retrattile a forbice Gorter %s (cassonetto %dx%dx%d mm)" % (
            stair,
            cw,
            cl,
            CASE_H,
        )
    return (
        "predisposta per scala retrattile a forbice (cassonetto %dx%dx%d mm, tipo Gorter %s)"
        % (
            cw,
            cl,
            CASE_H,
            stair,
        )
    )


def type_values(i):
    model, w, l = TYPES[i]
    area = ("%.2f" % (w * l / 1e6)).replace(".", ",")
    return [
        (BuiltInParameter.ALL_MODEL_TYPE_MARK, mark(i)),
        (BuiltInParameter.KEYNOTE_PARAM, "AR_%s_%s" % (CAT, mark(i))),
        (BuiltInParameter.ALL_MODEL_TYPE_COMMENTS, "BOTOLA %dx%d" % (w, l)),  # tag text
        (
            BuiltInParameter.ALL_MODEL_DESCRIPTION,
            "Botola di accesso alla copertura per manutenzione linea vita tipo Gorter %s o equivalente, "
            "opaca coibentata in alluminio a taglio termico, Uw 0,276 W/m2K, apertura controbilanciata "
            "con blocco in posizione aperta, guarnizione perimetrale, converse in lamiera con risvolto sul "
            "telaio, controtelaio nello spessore della copertura, %s, luce netta %dx%d mm (%s m2) conforme "
            "al Decreto Regione Lombardia 119/2009."
            % (model, stair_text(i), w, l, area),
        ),
        (BuiltInParameter.ALL_MODEL_MANUFACTURER, "Gorter"),
        (BuiltInParameter.ALL_MODEL_MODEL, model),
    ]


# names checked with the same rules as the Model Checker (05.2); sizes checked against the decree
assert R.RE_FAMILY.match(FAMILY_NAME)
for _i, (_m, _w, _l) in enumerate(TYPES):
    assert R.RE_TYPE.match(type_name(_i))
    assert _w * _l / 1e6 >= MIN_AREA and min(_w, _l) >= MIN_SIDE
    assert (
        STAIRS[_i][1] <= _w and STAIRS[_i][2] <= _l
    )  # the casing fits under the clear opening


def set_type_values(fm, i, notes):
    for bip, value in type_values(i):
        fp = fm.get_Parameter(bip)
        if fp is None:
            if i == 0:
                notes.append("Parametro di tipo non presente nel template: %s" % bip)
        else:
            fm.Set(fp, value)


def fixed_extrusion(fdoc, plane, loops, start_mm, end_mm, solid=True):
    """Extrusion with fixed extents (not driven by parameters). It is born from 0 to a positive end, then
    start is moved first and end second, so start < end at every step (end_mm may be 0).
    """
    arr = CurveArrArray()
    for loop in loops:
        arr.Append(loop)
    ext = fdoc.FamilyCreate.NewExtrusion(solid, arr, plane, ft(max(end_mm, 1)))
    ext.get_Parameter(BuiltInParameter.EXTRUSION_START_PARAM).Set(ft(start_mm))
    ext.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM).Set(ft(end_mm))
    return ext


def dashed_pattern(fdoc):
    for name in ("Dash", "Tratteggio"):
        pattern = LinePatternElement.GetLinePatternElementByName(fdoc, name)
        if pattern is not None:
            return pattern.Id
    lp = LinePattern("Tratteggio area di manovra")
    segments = List[LinePatternSegment]()
    segments.Add(LinePatternSegment(LinePatternSegmentType.Dash, ft(3)))  # paper length
    segments.Add(LinePatternSegment(LinePatternSegmentType.Space, ft(2)))
    lp.SetSegments(segments)
    return LinePatternElement.Create(fdoc, lp).Id


def area_subcategory(fdoc):
    """Dashed grey subcategory for the manoeuvre area (plan projection)."""
    parent = fdoc.OwnerFamily.FamilyCategory
    sub = fdoc.Settings.Categories.NewSubcategory(parent, SUB_AREA)
    sub.SetLinePatternId(dashed_pattern(fdoc), GraphicsStyleType.Projection)
    sub.LineColor = Color(128, 128, 128)
    return sub


def flex_test(fdoc, parts):
    """Every type made current in turn: each part must follow the clear opening. Rolled back."""
    fm = fdoc.FamilyManager
    types = dict((t.Name, t) for t in fm.Types)
    st = SubTransaction(fdoc)
    st.Start()
    try:
        results = []
        for i, (_model, w, l) in enumerate(TYPES):
            fm.CurrentType = types[type_name(i)]
            fdoc.Regenerate()
            for what, element, offset in parts:
                if offset is None:  # stair casing: its own size per type
                    expected = (STAIRS[i][2], STAIRS[i][1])
                else:
                    expected = (l + 2 * offset, w + 2 * offset)
                got = size_mm(element)
                results.append(
                    ("%s %s" % (mark(i), what), got, expected, same(got, expected))
                )
    except Exception as ex:
        results = [("Flessione", str(ex), "-", False)]
    finally:
        st.RollBack()
    return results


def build(fdoc, set_category):
    """Everything inside the family (open transaction). Returns (report rows, flex results, notes)."""
    fm = fdoc.FamilyManager
    notes, step = [], ""
    try:
        step = "categoria"
        if set_category:
            fdoc.OwnerFamily.FamilyCategory = Category.GetCategory(fdoc, BIC)
            fdoc.Regenerate()

        step = "tipo e parametri"
        _model, w, l = TYPES[0]
        if fm.Types.Size == 0:
            fm.NewType(type_name(0))
        else:
            fm.RenameCurrentType(type_name(0))
        p_w = length_param(fm, P_W, w)
        p_l = length_param(fm, P_L, l)
        p_prof = length_param(fm, P_PROF, PROFILE)
        p_flash_t = length_param(fm, P_FLASH_T, FLASH_T)
        p_flash_w = length_param(fm, P_FLASH_W, FLASH_W)
        p_flash_h = length_param(fm, P_FLASH_H, FLASH_H)
        p_over = length_param(fm, P_OVER, LID_OVER)
        p_area = length_param(fm, P_AREA, AREA)
        p_case_w = length_param(fm, P_CASE_W, STAIRS[0][1])
        p_case_l = length_param(fm, P_CASE_L, STAIRS[0][2])
        length_param(fm, P_W_OUT, formula="%s + 2 * %s" % (P_W, P_PROF))
        length_param(fm, P_L_OUT, formula="%s + 2 * %s" % (P_L, P_PROF))
        length_param(fm, P_SURF, formula="%s * %s" % (P_W, P_L), spec=SpecTypeId.Area)
        p_curb = length_param(fm, P_CURB, CURB)
        length_param(fm, P_GASKET, GASKET)
        length_param(fm, P_LID, LID)
        z_lid_base = length_param(
            fm, Z_LID_BASE, formula="%s + %s" % (P_CURB, P_GASKET)
        )
        z_lid = length_param(fm, Z_LID, formula="%s + %s" % (Z_LID_BASE, P_LID))
        p_show = fm.AddParameter(
            P_SHOW_AREA, GroupTypeId.Graphics, SpecTypeId.Boolean.YesNo, True
        )
        fm.Set(p_show, 1)
        p_mats = dict(
            (
                name,
                fm.AddParameter(
                    name, GroupTypeId.Materials, SpecTypeId.Reference.Material, False
                ),
            )
            for name, _parts in MATERIALS
        )
        fdoc.Regenerate()

        step = "piani di riferimento"
        view = plan_view(fdoc)
        planes = Planes(fdoc, view, 4000)
        hl, hw = l / 2.0, w / 2.0
        offsets = [  # (plane prefix, offset from the clear opening)
            ("Luce", 0),
            ("Esterno", PROFILE),
            ("Risvolto", PROFILE + FLASH_T),
            ("Converse", PROFILE + FLASH_W),
            ("Coperchio", PROFILE + LID_OVER),
            ("Manovra", PROFILE + AREA),
        ]
        for prefix, off in offsets:
            planes.add(prefix + " sinistra", "x", -hl - off)
            planes.add(prefix + " destra", "x", hl + off)
            planes.add(prefix + " fronte", "y", -hw - off)
            planes.add(prefix + " retro", "y", hw + off)
        # stair casing: own planes, may coincide with the clear opening in the first type (Small+)
        cw, cl = STAIRS[0][1], STAIRS[0][2]
        case_planes = [
            "Cassonetto " + s for s in ("sinistra", "destra", "fronte", "retro")
        ]
        for name, axis, at in zip(
            case_planes, "xxyy", (-cl / 2.0, cl / 2.0, -cw / 2.0, cw / 2.0)
        ):
            planes.add(name, axis, at)

        step = "estrusioni"
        plane = level_plane(fdoc)

        def ring(inner_off, outer_off):
            return [
                rectangle(l + 2 * outer_off, w + 2 * outer_off),
                rectangle(l + 2 * inner_off, w + 2 * inner_off),
            ]

        outer = PROFILE
        parts = {}
        parts["Controtelaio"] = fixed_extrusion(
            fdoc, plane, ring(0, outer), -HOST_THK, 0
        )
        parts["Cassonetto scala"] = fixed_extrusion(
            fdoc, plane, [rectangle(cl, cw)], -HOST_THK - CASE_H, -HOST_THK
        )
        parts["Telaio"] = extrude(fdoc, plane, ring(0, outer), None, p_curb)
        parts["Converse"] = extrude(
            fdoc, plane, ring(outer, PROFILE + FLASH_W), None, p_flash_t
        )
        parts["Risvolto converse"] = extrude(
            fdoc, plane, ring(outer, PROFILE + FLASH_T), p_flash_t, p_flash_h
        )
        parts["Guarnizione"] = extrude(fdoc, plane, ring(0, outer), p_curb, z_lid_base)
        parts["Coperchio"] = extrude(
            fdoc,
            plane,
            [rectangle(l + 2 * (PROFILE + LID_OVER), w + 2 * (PROFILE + LID_OVER))],
            z_lid_base,
            z_lid,
        )
        area = fixed_extrusion(
            fdoc, plane, ring(PROFILE + FLASH_W, PROFILE + AREA), 0, 1
        )
        parts["Area di manovra"] = area
        hole = fixed_extrusion(
            fdoc,
            plane,
            [rectangle(l + 2 * outer, w + 2 * outer)],
            -HOLE_DEPTH,
            HOLE_MARGIN,
            solid=False,
        )
        parts["Foro"] = hole

        step = "area di manovra (sottocategoria e visibilita')"
        area.Subcategory = area_subcategory(fdoc)
        vis = FamilyElementVisibility(FamilyElementVisibilityType.Model)
        vis.IsShownInFrontBack = False
        vis.IsShownInLeftRight = False
        area.SetVisibility(vis)
        fm.AssociateElementParameterToFamilyParameter(
            area.get_Parameter(BuiltInParameter.IS_VISIBLE_PARAM), p_show
        )

        step = "materiali"
        for name, names in MATERIALS:
            for part in names:
                fm.AssociateElementParameterToFamilyParameter(
                    parts[part].get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM),
                    p_mats[name],
                )
        fdoc.Regenerate()

        step = "vincoli delle facce"
        others = [n for n in planes.rp if n not in case_planes]
        locks = sum(
            planes.lock_faces(
                e, what, case_planes if what == "Cassonetto scala" else others
            )
            for what, e in sorted(parts.items())
        )
        notes += planes.failed

        step = "quote con etichetta"
        planes.dimension(
            ["Luce sinistra", "Luce destra"], -hw - PROFILE - AREA - 600, p_l
        )
        planes.dimension(
            ["Luce sinistra", "centro x", "Luce destra"],
            -hw - PROFILE - AREA - 300,
            equal=True,
        )
        planes.dimension(["Luce fronte", "Luce retro"], -hl - PROFILE - AREA - 600, p_w)
        planes.dimension(
            ["Luce fronte", "centro y", "Luce retro"],
            -hl - PROFILE - AREA - 300,
            equal=True,
        )
        planes.dimension(case_planes[:2], -hw - PROFILE - AREA - 1200, p_case_l)
        planes.dimension(
            [case_planes[0], "centro x", case_planes[1]],
            -hw - PROFILE - AREA - 900,
            equal=True,
        )
        planes.dimension(case_planes[2:], -hl - PROFILE - AREA - 1200, p_case_w)
        planes.dimension(
            [case_planes[2], "centro y", case_planes[3]],
            -hl - PROFILE - AREA - 900,
            equal=True,
        )
        for k, (prefix, label) in enumerate(
            [
                ("Esterno", p_prof),
                ("Risvolto", p_flash_t),
                ("Converse", p_flash_w),
                ("Coperchio", p_over),
                ("Manovra", p_area),
            ]
        ):
            base = "Luce" if prefix == "Esterno" else "Esterno"
            at_y, at_x = (
                hw + PROFILE + AREA + 200 + 150 * k,
                hl + PROFILE + AREA + 200 + 150 * k,
            )
            planes.dimension([prefix + " sinistra", base + " sinistra"], at_y, label)
            planes.dimension([base + " destra", prefix + " destra"], at_y, label)
            planes.dimension([prefix + " fronte", base + " fronte"], at_x, label)
            planes.dimension([base + " retro", prefix + " retro"], at_x, label)

        step = "tipi"
        first = fm.CurrentType
        set_type_values(fm, 0, notes)
        for i in range(1, len(TYPES)):
            fm.NewType(type_name(i))  # copies the current values, becomes current
            fm.Set(p_w, ft(TYPES[i][1]))
            fm.Set(p_l, ft(TYPES[i][2]))
            fm.Set(p_case_w, ft(STAIRS[i][1]))
            fm.Set(p_case_l, ft(STAIRS[i][2]))
            set_type_values(fm, i, notes)
        fm.CurrentType = first
        fdoc.Regenerate()

        step = "prova di flessione"
        flex = flex_test(
            fdoc,
            [
                ("Controtelaio", parts["Controtelaio"], PROFILE),
                ("Telaio", parts["Telaio"], PROFILE),
                ("Guarnizione", parts["Guarnizione"], PROFILE),
                ("Converse", parts["Converse"], PROFILE + FLASH_W),
                ("Risvolto", parts["Risvolto converse"], PROFILE + FLASH_T),
                ("Coperchio", parts["Coperchio"], PROFILE + LID_OVER),
                ("Area di manovra", area, PROFILE + AREA),
                ("Foro", hole, PROFILE),
                ("Cassonetto scala", parts["Cassonetto scala"], None),
            ],
        )
    except Exception as ex:
        raise Exception("%s: %s" % (step, ex))
    rows = [
        [
            "Piani di riferimento",
            "%d nuovi + 2 centrali del template" % (len(planes.rp) - 2),
        ],
        ["Vincoli delle facce", str(locks)],
        [
            "Quote con etichetta",
            "%s, %s + 2 EQ; %s, %s, %s, %s, %s (x4)"
            % (P_L, P_W, P_PROF, P_FLASH_T, P_FLASH_W, P_OVER, P_AREA),
        ],
        ["Materiali (parametri di tipo)", ", ".join(name for name, _p in MATERIALS)],
    ]
    return rows, flex, notes


# ---------------------------------------------------------------- checks (read only)

if not doc.IsFamilyDocument:
    forms.alert(
        "Apri prima il template '%s' (File > Nuovo > Famiglia) e lancia il comando dall'editor di famiglie."
        % TEMPLATE,
        exitscript=True,
    )
if doc.OwnerFamily.FamilyPlacementType != FamilyPlacementType.WorkPlaneBased:
    forms.alert(
        "La famiglia aperta non e' basata su superficie: serve il template '%s'. Nessuna modifica."
        % TEMPLATE,
        exitscript=True,
    )
category = doc.OwnerFamily.FamilyCategory
allowed = [
    Category.GetCategory(doc, b).Id for b in (BIC, BuiltInCategory.OST_GenericModel)
]
if category is None or category.Id not in allowed:
    forms.alert(
        "Categoria della famiglia: '%s'. Servono 'Modelli generici' (template '%s') o 'Attrezzature speciali': "
        "nessuna modifica." % (category.Name if category else "nessuna", TEMPLATE),
        exitscript=True,
    )
set_category = category.Id != Category.GetCategory(doc, BIC).Id
if doc.FamilyManager.get_Parameter(P_W):
    forms.alert(
        "La famiglia contiene gia' una botola: il comando lavora solo su un template appena aperto "
        "(File > Nuovo > Famiglia > %s). Nessuna modifica." % TEMPLATE,
        exitscript=True,
    )
if not forms.alert(
    "Template basato su superficie verificato. Categoria: %s%s.\n"
    "Creare la famiglia %s con %d tipi (%s)?"
    % (
        category.Name,
        " -> sara' impostata 'Attrezzature speciali'" if set_category else "",
        FAMILY_NAME,
        len(TYPES),
        ", ".join("%dx%d" % (w, l) for _m, w, l in TYPES),
    ),
    yes=True,
    no=True,
):
    script.exit()

# ---------------------------------------------------------------- build

t = Transaction(doc, "BMC - Crea botola copertura")
t.Start()
try:
    rows, flex, notes = build(doc, set_category)
except Exception as ex:
    t.RollBack()
    forms.alert("Errore in '%s'\nFamiglia non modificata." % ex, exitscript=True)
status = str(t.Commit())
if status != "Committed":
    forms.alert("Transazione %s: famiglia non modificata." % status, exitscript=True)

# ---------------------------------------------------------------- save, report, cut geometry

rfa = forms.save_file(
    file_ext="rfa", default_name=FAMILY_NAME, title="Salva la famiglia con il nome PGI"
)
if rfa:
    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True  # the save dialog already asked about overwriting
    doc.SaveAs(rfa, opts)

output.print_md("# BMC - Crea botola copertura")
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
        ["Categoria", "Attrezzature speciali (%s), basata su superficie" % CAT],
        ["File", rfa or "-"],
        [
            "Controtelaio",
            "profilo %d mm nello spessore dell'ospite (%d mm, fisso)"
            % (PROFILE, HOST_THK),
        ],
        [
            "Telaio",
            "profilo %d mm, '%s' = %d mm sopra l'ospite" % (PROFILE, P_CURB, CURB),
        ],
        [
            "Converse",
            "lamiera %d mm larga %d mm, risvolto alto %d mm"
            % (FLASH_T, FLASH_W, FLASH_H),
        ],
        [
            "Guarnizione / coperchio",
            "%d mm / %d mm, sporgenza %d mm" % (GASKET, LID, LID_OVER),
        ],
        [
            "Area di manovra",
            "%d mm oltre il telaio, tratteggiata, solo in pianta ('%s', istanza)"
            % (AREA, P_SHOW_AREA),
        ],
        [
            "Foro nell'ospite",
            "vuoto grande quanto il telaio, da -%d a +%d mm"
            % (HOLE_DEPTH, HOLE_MARGIN),
        ],
        [
            "Cassonetto scala",
            "scala a forbice chiusa, profondita' %d mm sotto il controtelaio (fisso), "
            "dimensioni per tipo in '%s' / '%s'" % (CASE_H, P_CASE_W, P_CASE_L),
        ],
    ]
    + rows,
    columns=["Voce", "Valore"],
)
output.print_md(
    "## Tipi (Decreto Regione Lombardia 119/2009: superficie >= 0,50 m2, lato minore >= 0,70 m)"
)
output.print_table(
    table_data=[
        [
            type_name(i),
            model,
            "%dx%d" % (w, l),
            "%.2f" % (w * l / 1e6),
            "si",
            "%s %dx%d%s"
            % (
                STAIRS[i][0],
                STAIRS[i][1],
                STAIRS[i][2],
                "" if STAIRS[i][3] else " (predisposizione, da confermare)",
            ),
        ]
        for i, (model, w, l) in enumerate(TYPES)
    ],
    columns=[
        "Tipo",
        "Gorter",
        "Luce netta (mm)",
        "Superficie (m2)",
        "Conforme",
        "Scala / cassonetto",
    ],
)
output.print_md("## Prova di flessione (ogni tipo reso corrente, poi annullata)")
bad = [f for f in flex if not f[3]]
output.print_md(
    "**%d controlli su %d OK.**" % (len(flex) - len(bad), len(flex))
    if not bad
    else "**ATTENZIONE: %d controlli su %d NON seguono i tipi:**"
    % (len(bad), len(flex))
)
if bad:
    output.print_table(
        table_data=[
            [
                name,
                " x ".join("%.0f" % g for g in got) if isinstance(got, tuple) else got,
                " x ".join("%.0f" % e for e in exp) if isinstance(exp, tuple) else exp,
            ]
            for name, got, exp, _ok in bad
        ],
        columns=["Elemento", "Ottenuto (mm)", "Atteso (mm)"],
    )
for note in notes:
    output.print_md("- %s" % note)
output.print_md(
    "**Da verificare sul disegno Gorter RHT:** profilo telaio, altezza telaio, guarnizione, coperchio "
    "e sporgenza, converse (parametri di tipo). Materiali: assegnali nei parametri 'Materiale ...'."
)
output.print_md(
    "## Posizionamento (docs/20260928/Botola_normativa_posizionamento.md)\n"
    "- area di manovra libera: comignoli, canne e ostacoli fuori dalla lastra tratteggiata;\n"
    "- area di manovra tutta a >= 2,3 m dalla gronda, altrimenti protezione collettiva sul bordo;\n"
    "- primo ancoraggio UNI EN 795 dentro l'area di manovra, a ~0,60 m dal telaio;\n"
    "- disattiva '%s' nelle viste 3D o prima dell'export IFC." % P_SHOW_AREA
)
output.print_md(
    "## Etichetta per le tavole (manuale: l'API non crea etichette nelle famiglie di etichette)\n"
    "File > Nuovo > Simbolo di annotazione > `Metric Specialty Equipment Tag.rft`: etichetta con "
    "**Contrassegno tipo** e **Commenti tipo** (es. `BT1.01A` / `BOTOLA 700x1000`), salva come "
    "`AR_TAG_Attrezzature speciali.rfa`, caricala e usala in pianta copertura e in sezione."
)
output.print_md(
    "## Ultimo passo manuale: foro nell'ospite\n"
    "Si e' avviato **Taglia geometria**: in una vista 3D o in un prospetto clicca prima il **finto ospite** "
    "grigio del template, poi il **vuoto**. Poi **Ctrl+S**."
)
__revit__.PostCommand(
    RevitCommandId.LookupPostableCommandId(PostableCommand.CutGeometry)
)
