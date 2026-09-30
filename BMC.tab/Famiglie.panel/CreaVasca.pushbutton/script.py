"""Build the parametric Pircher V.A. 30 precast tank inside the open Specialty Equipment family template"""

__title__ = "Crea vasca\nVA30"
__author__ = "Luca Rosati"
__guida__ = "Da lanciare nell'editor di famiglie, sul template Attrezzature speciali appena aperto: verifica la categoria e costruisce la vasca Pircher V.A. 30 parametrica (piani di riferimento, quote con etichetta, pareti, fondo, piastra e 2 chiusini nidificati), prova la flessione e propone il salvataggio con il nome PGI. Il caricamento nel progetto resta manuale."

# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# PGI R06 3.11.5-3.11.6: family AR_ATS_VA1.00A_<descr>, type AR_ATS_VA1.01A_<descr>_<LxWxH mm>, Type Mark VA1.01A.
# Dimensions: Pircher catalogue, "Vasca rettangolare" table, row v.a. 30 (b x d x f external, a x c x e internal).
# The cover slab sits on the walls (F excludes it: 5.80 x 2.30 x 2.35 = 30 m3). Slab, collar and manhole sizes
# are not in the table: defaults read on the drawing or assumed, to be checked on the Pircher slab datasheet.
#
# Parametric plan: reference planes (outer, inner, manhole axes) driven by labelled dimensions (EQ about the
# template centre planes); the vertical faces of walls, bottom and slab are locked to them. A circle has no
# planar face to lock, so each manhole (collar + lid + void through the slab) is a nested Generic Model family
# whose centre planes are locked to the manhole axis planes; its void cuts the slab.
# Heights: extrusion start/end bound to family parameters (formulas for the levels).
# ponytail: manhole diameter fixed (MH_D); make it a nested parameter with a labelled radial dimension if the
# 80 cm manholes are needed.
# Runs on a freshly opened template: the nested family is loaded first (LoadFamily needs no open transaction),
# then one Transaction builds everything (Ctrl+Z undoes it). A flex test in a SubTransaction checks that the
# geometry follows length and width, then is rolled back.

import math
import os
import tempfile

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType

from pyrevit import forms
from pyrevit import script

import bmc_rules as R

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
output = script.get_output()

BIC = BuiltInCategory.OST_SpecialityEquipment
CAT = R.CAT_CODE[BIC]  # ATS
TEMPLATE = "Metric Specialty Equipment.rft"
NESTED_TEMPLATE = "Metric Generic Model.rft"

# catalogue row v.a. 30, mm
LENGTH, WIDTH, HEIGHT = 6000, 2500, 2500  # b, d, f
WALL, BOTTOM = 100, 150  # (b - a) / 2 = (d - c) / 2, f - e
# cover, mm: not in the catalogue table (see header)
SLAB, COLLAR, LID = (
    200,
    150,
    50,
)  # slab thickness, collar height above the slab, lid thickness
MH_D = 600  # manhole clear diameter (specification: "chiusini in ghisa D 60 e/o 80 cm")
MH_RING = 100  # collar wall thickness
MH_OFFSET = (
    650  # manhole centre from the short outer side, measured on the plan drawing
)
HOLE_MARGIN = 10  # the void overshoots the slab faces: no coplanar cut
FLEX = (1000, 500)  # flex test: length and width increments

MARK = "VA1.01A"
FAMILY_NAME = "AR_%s_VA1.00A_Vasca monolitica interrata" % CAT
TYPE_NAME = "AR_%s_%s_Vasca monolitica VA30_%dx%dx%d" % (
    CAT,
    MARK,
    LENGTH,
    WIDTH,
    HEIGHT,
)
MH_FAMILY = (
    "Chiusino vasca D%d" % MH_D
)  # nested, not shared: never listed in the project
TYPE_VALUES = [
    (BuiltInParameter.ALL_MODEL_TYPE_MARK, MARK),
    (BuiltInParameter.KEYNOTE_PARAM, "AR_%s_%s" % (CAT, MARK)),
    (
        BuiltInParameter.ALL_MODEL_DESCRIPTION,
        "Vasca monolitica parallelepipeda in calcestruzzo C40/50 Pircher V.A. 30, cod. art. 7630, "
        "volume 30000 l, peso 17,0 t, grezza. Dimensioni esterne 6000x2500x2500 mm, "
        "interne 5800x2300x2350 mm. Piastra carrabile di copertura con 2 chiusini in ghisa "
        "diametro %d mm su rialzo." % MH_D,
    ),
    (BuiltInParameter.ALL_MODEL_MANUFACTURER, "Pircher"),
    (BuiltInParameter.ALL_MODEL_MODEL, "V.A. 30 cod. 7630"),
]

# host parameters (type)
P_LEN, P_WID, P_WALL, P_MH = (
    "Lunghezza esterna",
    "Larghezza esterna",
    "Spessore pareti",
    "Distanza chiusini",
)
P_LEN_IN, P_WID_IN = "Lunghezza interna", "Larghezza interna"
P_HEIGHT, P_BOTTOM, P_INNER = "Altezza esterna", "Spessore fondo", "Altezza interna"
P_SLAB, P_COLLAR, P_LID = (
    "Spessore piastra",
    "Altezza rialzo chiusini",
    "Spessore chiusino",
)
Z_SLAB = "Quota piastra"
# nested manhole parameters (instance); P_SLAB, P_COLLAR, P_LID keep the host names
N_BASE, N_TOP, N_LID = "Quota base", "Quota rialzo", "Quota chiusino"
N_MARGIN, N_HOLE_BOT, N_HOLE_TOP = (
    "Margine foro",
    "Quota fondo foro",
    "Quota testa foro",
)

# names checked with the same rules as the Model Checker (05.2)
assert R.RE_FAMILY.match(FAMILY_NAME) and R.RE_TYPE.match(TYPE_NAME)


def ft(mm):
    return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters)


def to_mm(value_ft):
    return UnitUtils.ConvertFromInternalUnits(value_ft, UnitTypeId.Millimeters)


def template_path(name):
    folders = [
        app.FamilyTemplatePath,
        os.path.join(
            os.environ.get("ProgramData", r"C:\ProgramData"),
            "Autodesk",
            "RVT %s" % app.VersionNumber,
            "Family Templates",
            "English",
        ),
    ]
    for folder in folders:
        path = os.path.join(folder or "", name)
        if os.path.exists(path):
            return path
    path = forms.pick_file(file_ext="rft", title="Scegli il template '%s'" % name)
    if not path:
        script.exit()
    return path


# ---------------------------------------------------------------- geometry helpers


def rectangle(length, width):
    x, y = ft(length) / 2.0, ft(width) / 2.0
    pts = [XYZ(-x, -y, 0), XYZ(x, -y, 0), XYZ(x, y, 0), XYZ(-x, y, 0)]
    loop = CurveArray()
    for i in range(4):
        loop.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
    return loop


def circle(diameter):
    """Closed loop centred on the origin, two half arcs (a single full arc is not a reliable sketch profile)."""
    loop = CurveArray()
    for start in (0.0, math.pi):
        loop.Append(
            Arc.Create(
                XYZ.Zero,
                ft(diameter) / 2.0,
                start,
                start + math.pi,
                XYZ.BasisX,
                XYZ.BasisY,
            )
        )
    return loop


def level_plane(fdoc):
    levels = list(FilteredElementCollector(fdoc).OfClass(Level))
    if levels:
        return SketchPlane.Create(fdoc, levels[0].Id)
    return SketchPlane.Create(fdoc, Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero))


def length_param(fm, name, value_mm=None, formula=None, instance=False):
    p = fm.AddParameter(name, GroupTypeId.Geometry, SpecTypeId.Length, instance)
    if formula:
        fm.SetFormula(p, formula)
    else:
        fm.Set(p, ft(value_mm))
    return p


def extrude(fdoc, plane, loops, start, end, solid=True):
    """Extrusion bound to family parameters: end first, then start (start < end at every step).
    start None = fixed at 0."""
    fm = fdoc.FamilyManager
    arr = CurveArrArray()
    for loop in loops:
        arr.Append(loop)
    ext = fdoc.FamilyCreate.NewExtrusion(
        solid, arr, plane, fm.CurrentType.AsDouble(end)
    )
    fm.AssociateElementParameterToFamilyParameter(
        ext.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM), end
    )
    if start is None:
        ext.get_Parameter(BuiltInParameter.EXTRUSION_START_PARAM).Set(0.0)
    else:
        fm.AssociateElementParameterToFamilyParameter(
            ext.get_Parameter(BuiltInParameter.EXTRUSION_START_PARAM), start
        )
    return ext


# ---------------------------------------------------------------- nested manhole family


def build_manhole(ndoc):
    """Collar ring + lid + void through the slab, all around the origin; heights from instance parameters
    that the host drives. Runs inside a transaction of the nested document."""
    fm = ndoc.FamilyManager
    if fm.Types.Size == 0:
        fm.NewType(MH_FAMILY)
    ndoc.OwnerFamily.get_Parameter(BuiltInParameter.FAMILY_ALLOW_CUT_WITH_VOIDS).Set(1)
    length_param(fm, N_BASE, HEIGHT + SLAB, instance=True)
    length_param(fm, P_SLAB, SLAB, instance=True)
    length_param(fm, P_COLLAR, COLLAR, instance=True)
    length_param(fm, P_LID, LID, instance=True)
    length_param(fm, N_MARGIN, HOLE_MARGIN, instance=True)
    top = length_param(fm, N_TOP, formula="%s + %s" % (N_BASE, P_COLLAR), instance=True)
    lid = length_param(fm, N_LID, formula="%s - %s" % (N_TOP, P_LID), instance=True)
    hole_bot = length_param(
        fm,
        N_HOLE_BOT,
        formula="%s - %s - %s" % (N_BASE, P_SLAB, N_MARGIN),
        instance=True,
    )
    hole_top = length_param(
        fm, N_HOLE_TOP, formula="%s + %s" % (N_BASE, N_MARGIN), instance=True
    )
    ndoc.Regenerate()
    plane = level_plane(ndoc)
    base = fm.get_Parameter(N_BASE)
    extrude(ndoc, plane, [circle(MH_D + 2 * MH_RING), circle(MH_D)], base, top)
    extrude(ndoc, plane, [circle(MH_D)], lid, top)
    extrude(ndoc, plane, [circle(MH_D)], hole_bot, hole_top, solid=False)


def manhole_family(fdoc):
    """The nested family, loaded into fdoc (reused when a previous run already loaded it)."""
    for fam in FilteredElementCollector(fdoc).OfClass(Family):
        if fam.Name == MH_FAMILY:
            return fam
    ndoc = app.NewFamilyDocument(template_path(NESTED_TEMPLATE))
    try:
        t = Transaction(ndoc, "BMC - Chiusino")
        t.Start()
        try:
            build_manhole(ndoc)
        except Exception:
            t.RollBack()
            raise
        if str(t.Commit()) != "Committed":
            raise Exception("transazione del chiusino non confermata")
        path = os.path.join(
            tempfile.gettempdir(), MH_FAMILY + ".rfa"
        )  # the file name is the family name
        opts = SaveAsOptions()
        opts.OverwriteExistingFile = True
        ndoc.SaveAs(path, opts)
        return ndoc.LoadFamily(fdoc)
    finally:
        ndoc.Close(False)


# ---------------------------------------------------------------- host family


class Planes(object):
    """Reference planes of the plan: coordinate (ft) per name, planes normal to X ("x") or to Y ("y")."""

    def __init__(self, fdoc, view):
        self.fdoc, self.view, self.rp, self.axis, self.at = fdoc, view, {}, {}, {}
        self.failed = []
        # views where a lock can be made: the plan first, then the elevation seeing the planes edge-on
        # (the slab and the manholes are above the plan cut plane)
        elevations = [
            v
            for v in FilteredElementCollector(fdoc).OfClass(View)
            if not v.IsTemplate and v.ViewType == ViewType.Elevation
        ]
        self.views = {
            "x": [view] + [v for v in elevations if abs(v.ViewDirection.Y) > 0.99],
            "y": [view] + [v for v in elevations if abs(v.ViewDirection.X) > 0.99],
        }
        for rp in FilteredElementCollector(fdoc).OfClass(ReferencePlane):
            n = rp.Normal
            if abs(n.X) > 0.99 and abs(rp.BubbleEnd.X) < 1e-6:
                self._keep("centro x", rp, "x", 0.0)
            elif abs(n.Y) > 0.99 and abs(rp.BubbleEnd.Y) < 1e-6:
                self._keep("centro y", rp, "y", 0.0)
        if len(self.rp) != 2:
            raise Exception("piani centrali del template non trovati")

    def _keep(self, name, rp, axis, at):
        self.rp[name], self.axis[name], self.at[name] = rp, axis, at

    def add(self, name, axis, at_mm):
        at, big = ft(at_mm), ft(LENGTH)
        if axis == "x":
            a, b = XYZ(at, -big, 0), XYZ(at, big, 0)
        else:
            a, b = XYZ(-big, at, 0), XYZ(big, at, 0)
        rp = self.fdoc.FamilyCreate.NewReferencePlane(a, b, XYZ.BasisZ, self.view)
        rp.Name = name
        self._keep(name, rp, axis, at)

    def lock(self, name, reference, what):
        """Locked alignment of reference to plane name, tried in each view of its axis. Returns 1 or 0."""
        error = None
        for view in self.views[self.axis[name]]:
            try:
                self.fdoc.FamilyCreate.NewAlignment(
                    view, self.rp[name].GetReference(), reference
                )
                return 1
            except Exception as ex:
                error = ex
        self.failed.append("%s non vincolato a '%s': %s" % (what, name, error))
        return 0

    def lock_faces(self, element, what):
        """Lock every vertical planar face of element to the reference plane it lies on. Returns the count."""
        opts = Options()
        opts.ComputeReferences = True
        count = 0
        for geo in element.get_Geometry(opts):
            if not isinstance(geo, Solid):
                continue
            for face in geo.Faces:
                if not isinstance(face, PlanarFace) or face.Reference is None:
                    continue
                n = face.FaceNormal
                axis = "x" if abs(n.X) > 0.99 else "y" if abs(n.Y) > 0.99 else None
                if axis is None:
                    continue
                at = face.Origin.X if axis == "x" else face.Origin.Y
                for name in self.rp:
                    if self.axis[name] == axis and abs(self.at[name] - at) < 1e-4:
                        count += self.lock(name, face.Reference, what)
                        break
        return count

    def dimension(self, names, line_at_mm, label=None, equal=False):
        refs = ReferenceArray()
        for name in names:
            refs.Append(self.rp[name].GetReference())
        at, big = ft(line_at_mm), ft(LENGTH)
        if self.axis[names[0]] == "x":
            line = Line.CreateBound(XYZ(-big, at, 0), XYZ(big, at, 0))
        else:
            line = Line.CreateBound(XYZ(at, -big, 0), XYZ(at, big, 0))
        dim = self.fdoc.FamilyCreate.NewDimension(self.view, line, refs)
        if label is not None:
            dim.FamilyLabel = label
        if equal:
            dim.AreSegmentsEqual = True
        return dim


def plan_view(fdoc):
    for v in FilteredElementCollector(fdoc).OfClass(ViewPlan):
        if not v.IsTemplate and v.GenLevel is not None:
            return v
    raise Exception("vista di pianta del livello di riferimento non trovata")


def size_mm(element):
    bb = element.get_BoundingBox(None)
    return to_mm(bb.Max.X - bb.Min.X), to_mm(bb.Max.Y - bb.Min.Y)


def flex_test(fdoc, p_len, p_wid, walls, bottom, slab, manholes):
    """Change length and width, check the geometry, roll back. Returns [(item, got, expected, ok)]."""
    fm = fdoc.FamilyManager
    length, width = LENGTH + FLEX[0], WIDTH + FLEX[1]
    st = SubTransaction(fdoc)
    st.Start()
    try:
        fm.Set(p_len, ft(length))
        fm.Set(p_wid, ft(width))
        fdoc.Regenerate()
        checks = [
            ("Pareti", size_mm(walls), (length, width)),
            ("Fondo", size_mm(bottom), (length - 2 * WALL, width - 2 * WALL)),
            ("Piastra", size_mm(slab), (length, width)),
        ]
        for inst in manholes:
            p = inst.Location.Point
            checks.append(
                (
                    "Chiusino",
                    (to_mm(abs(p.X)), to_mm(abs(p.Y))),
                    (length / 2.0 - MH_OFFSET, 0),
                )
            )
        results = [
            (name, got, exp, all(abs(g - e) < 1 for g, e in zip(got, exp)))
            for name, got, exp in checks
        ]
    except Exception as ex:
        results = [("Flessione", str(ex), "-", False)]
    finally:
        st.RollBack()
    return results


def build(fdoc, mh_family):
    """Everything inside the host family (open transaction). Returns (report rows, flex results, notes)."""
    fm = fdoc.FamilyManager
    notes, step = [], ""
    try:
        step = "tipo e parametri"
        if fm.Types.Size == 0:
            fm.NewType(TYPE_NAME)
        else:
            fm.RenameCurrentType(TYPE_NAME)
        p_len = length_param(fm, P_LEN, LENGTH)
        p_wid = length_param(fm, P_WID, WIDTH)
        p_wall = length_param(fm, P_WALL, WALL)
        p_mh = length_param(fm, P_MH, MH_OFFSET)
        length_param(fm, P_LEN_IN, formula="%s - 2 * %s" % (P_LEN, P_WALL))
        length_param(fm, P_WID_IN, formula="%s - 2 * %s" % (P_WID, P_WALL))
        p_height = length_param(fm, P_HEIGHT, HEIGHT)
        p_bottom = length_param(fm, P_BOTTOM, BOTTOM)
        length_param(fm, P_INNER, formula="%s - %s" % (P_HEIGHT, P_BOTTOM))
        p_slab = length_param(fm, P_SLAB, SLAB)
        p_collar = length_param(fm, P_COLLAR, COLLAR)
        p_lid = length_param(fm, P_LID, LID)
        z_slab = length_param(fm, Z_SLAB, formula="%s + %s" % (P_HEIGHT, P_SLAB))
        fdoc.Regenerate()

        step = "piani di riferimento"
        view = plan_view(fdoc)
        planes = Planes(fdoc, view)
        half_l, half_w = LENGTH / 2.0, WIDTH / 2.0
        for name, axis, at in [
            ("Esterno sinistra", "x", -half_l),
            ("Esterno destra", "x", half_l),
            ("Interno sinistra", "x", -half_l + WALL),
            ("Interno destra", "x", half_l - WALL),
            ("Chiusino sinistra", "x", -half_l + MH_OFFSET),
            ("Chiusino destra", "x", half_l - MH_OFFSET),
            ("Esterno fronte", "y", -half_w),
            ("Esterno retro", "y", half_w),
            ("Interno fronte", "y", -half_w + WALL),
            ("Interno retro", "y", half_w - WALL),
        ]:
            planes.add(name, axis, at)

        step = "estrusioni"
        plane = level_plane(fdoc)
        inner = (LENGTH - 2 * WALL, WIDTH - 2 * WALL)
        walls = extrude(
            fdoc, plane, [rectangle(LENGTH, WIDTH), rectangle(*inner)], None, p_height
        )
        bottom = extrude(fdoc, plane, [rectangle(*inner)], None, p_bottom)
        slab = extrude(fdoc, plane, [rectangle(LENGTH, WIDTH)], p_height, z_slab)

        step = "chiusini nidificati"
        symbol = fdoc.GetElement(list(mh_family.GetFamilySymbolIds())[0])
        if not symbol.IsActive:
            symbol.Activate()
        level = list(FilteredElementCollector(fdoc).OfClass(Level))[0]
        manholes = []
        for x in (-(half_l - MH_OFFSET), half_l - MH_OFFSET):
            inst = fdoc.FamilyCreate.NewFamilyInstance(
                XYZ(ft(x), 0, 0), symbol, level, StructuralType.NonStructural
            )
            for nested, host in (
                (N_BASE, z_slab),
                (P_SLAB, p_slab),
                (P_COLLAR, p_collar),
                (P_LID, p_lid),
            ):
                fm.AssociateElementParameterToFamilyParameter(
                    inst.LookupParameter(nested), host
                )
            manholes.append(inst)
        fdoc.Regenerate()

        step = "foratura della piastra"
        cuts = 0
        for inst in manholes:
            try:
                InstanceVoidCutUtils.AddInstanceVoidCut(fdoc, slab, inst)
                cuts += 1
            except Exception as ex:
                notes.append("Foro nella piastra non creato: %s" % ex)
        fdoc.Regenerate()

        step = "vincoli delle facce"
        locks = sum(
            planes.lock_faces(e, what)
            for e, what in ((walls, "Pareti"), (bottom, "Fondo"), (slab, "Piastra"))
        )
        for inst, axis_plane in zip(manholes, ("Chiusino sinistra", "Chiusino destra")):
            for kind, target in (
                (FamilyInstanceReferenceType.CenterLeftRight, axis_plane),
                (FamilyInstanceReferenceType.CenterFrontBack, "centro y"),
            ):
                refs = inst.GetReferences(kind)
                if refs.Count == 0:
                    planes.failed.append(
                        "Chiusino senza riferimento %s: non vincolato a '%s'"
                        % (kind, target)
                    )
                    continue
                locks += planes.lock(target, refs[0], "Chiusino")
        notes += planes.failed

        step = "quote con etichetta"
        y_out, x_out = -half_w - 1500, -half_l - 1500
        planes.dimension(["Esterno sinistra", "Esterno destra"], y_out, p_len)
        planes.dimension(
            ["Esterno sinistra", "centro x", "Esterno destra"], y_out + 500, equal=True
        )
        planes.dimension(["Esterno fronte", "Esterno retro"], x_out, p_wid)
        planes.dimension(
            ["Esterno fronte", "centro y", "Esterno retro"], x_out + 500, equal=True
        )
        for pair in (
            ("Esterno sinistra", "Interno sinistra"),
            ("Interno destra", "Esterno destra"),
        ):
            planes.dimension(list(pair), half_w + 500, p_wall)
        for pair in (
            ("Esterno fronte", "Interno fronte"),
            ("Interno retro", "Esterno retro"),
        ):
            planes.dimension(list(pair), half_l + 500, p_wall)
        for pair in (
            ("Esterno sinistra", "Chiusino sinistra"),
            ("Chiusino destra", "Esterno destra"),
        ):
            planes.dimension(list(pair), half_w + 1000, p_mh)

        step = "dati del tipo"
        for bip, value in TYPE_VALUES:
            fp = fm.get_Parameter(bip)
            if fp is None:
                notes.append("Parametro di tipo non presente nel template: %s" % bip)
            else:
                fm.Set(fp, value)
        fdoc.Regenerate()

        step = "prova di flessione"
        flex = flex_test(fdoc, p_len, p_wid, walls, bottom, slab, manholes)
    except Exception as ex:
        raise Exception("%s: %s" % (step, ex))
    rows = [
        [
            "Piani di riferimento",
            "%d nuovi + 2 centrali del template" % (len(planes.rp) - 2),
        ],
        ["Vincoli (facce e chiusini)", str(locks)],
        [
            "Quote con etichetta",
            "%s, %s, %s (x4), %s (x2) + 2 EQ" % (P_LEN, P_WID, P_WALL, P_MH),
        ],
        ["Fori nella piastra", "%d di 2" % cuts],
    ]
    return rows, flex, notes


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
if doc.FamilyManager.get_Parameter(P_HEIGHT) or doc.FamilyManager.get_Parameter(P_LEN):
    forms.alert(
        "La famiglia contiene gia' una vasca: il comando lavora solo su un template appena aperto "
        "(File > Nuovo > Famiglia > %s). Nessuna modifica." % TEMPLATE,
        exitscript=True,
    )
if not forms.alert(
    "Categoria verificata: %s.\nCreare il tipo\n%s\nparametrico (piani, quote, pareti, fondo, piastra, 2 chiusini)?"
    % (category.Name, TYPE_NAME),
    yes=True,
    no=True,
):
    script.exit()

# ---------------------------------------------------------------- build

try:
    mh_family = manhole_family(doc)
except Exception as ex:
    forms.alert(
        "Famiglia del chiusino non creata, vasca non modificata:\n%s" % ex,
        exitscript=True,
    )

t = Transaction(doc, "BMC - Crea vasca VA30")
t.Start()
try:
    rows, flex, notes = build(doc, mh_family)
except Exception as ex:
    t.RollBack()
    forms.alert(
        "Errore in '%s'\nVasca non creata (resta caricata solo la famiglia '%s')."
        % (ex, MH_FAMILY),
        exitscript=True,
    )
status = str(t.Commit())
if status != "Committed":
    forms.alert("Transazione %s: vasca non creata." % status, exitscript=True)

# ---------------------------------------------------------------- save (optional) and report

rfa = forms.save_file(
    file_ext="rfa", default_name=FAMILY_NAME, title="Salva la famiglia con il nome PGI"
)
if rfa:
    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True  # the save dialog already asked about overwriting
    doc.SaveAs(rfa, opts)

output.print_md("# BMC - Crea vasca VA30")
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
        ["Tipo", TYPE_NAME],
        ["Categoria", "%s (%s)" % (category.Name, CAT)],
        ["File", rfa or "-"],
        [
            "Chiusini",
            "famiglia nidificata '%s', rialzo D%d/D%d"
            % (MH_FAMILY, MH_D + 2 * MH_RING, MH_D),
        ],
    ]
    + rows,
    columns=["Voce", "Valore"],
)
output.print_md(
    "## Prova di flessione (+%d mm lunghezza, +%d mm larghezza, poi annullata)" % FLEX
)
output.print_table(
    table_data=[
        [
            name,
            " x ".join("%.0f" % g for g in got) if isinstance(got, tuple) else got,
            " x ".join("%.0f" % e for e in exp) if isinstance(exp, tuple) else exp,
            "OK" if ok else "NO",
        ]
        for name, got, exp, ok in flex
    ],
    columns=["Elemento", "Ottenuto (mm)", "Atteso (mm)", "Esito"],
)
if all(f[3] for f in flex):
    output.print_md("**La famiglia segue lunghezza e larghezza.**")
else:
    output.print_md(
        "**ATTENZIONE: la geometria non segue tutti i parametri, vedi le righe NO.**"
    )
for note in notes:
    output.print_md("- %s" % note)
output.print_md(
    "**Da verificare sulla scheda Pircher della piastra:** spessore piastra, diametro chiusini (60 o 80 cm), "
    "altezza rialzi, posizione. Origine = centro del fondo. Poi *Carica nel progetto*."
)
