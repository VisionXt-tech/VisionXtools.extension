"""Helpers to build parametric loadable families from a pyRevit script run in the Family Editor.

Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
Extracted from Famiglie.panel/CreaVasca.pushbutton (verified in Revit 2024, 28/09/2026); CreaVasca still carries
its own copy and can move to this module the next time it is touched.
Rules behind these helpers: skill revit-family-builder (~/.claude/skills/revit-family-builder/SKILL.md).
"""

import math

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *


def ft(mm):
    return UnitUtils.ConvertToInternalUnits(mm, UnitTypeId.Millimeters)


def to_mm(value_ft):
    return UnitUtils.ConvertFromInternalUnits(value_ft, UnitTypeId.Millimeters)


def rectangle(length, width):
    """Closed loop centred on the origin, length along X, width along Y (mm)."""
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


def length_param(fm, name, value_mm=None, formula=None, instance=False, spec=None):
    """Family parameter (Dimensions group): a value in mm, or a formula. spec defaults to Length."""
    p = fm.AddParameter(name, GroupTypeId.Geometry, spec or SpecTypeId.Length, instance)
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


def plan_view(fdoc):
    for v in FilteredElementCollector(fdoc).OfClass(ViewPlan):
        if not v.IsTemplate and v.GenLevel is not None:
            return v
    raise Exception("vista di pianta del livello di riferimento non trovata")


def size_mm(element):
    bb = element.get_BoundingBox(None)
    return to_mm(bb.Max.X - bb.Min.X), to_mm(bb.Max.Y - bb.Min.Y)


def same(got, expected, tolerance_mm=1.0):
    return all(abs(g - e) < tolerance_mm for g, e in zip(got, expected))


class Planes(object):
    """Reference planes of the plan: coordinate (ft) per name, planes normal to X ("x") or to Y ("y").
    The template centre planes are found by normal and position ("centro x", "centro y"), not by name.
    """

    def __init__(self, fdoc, view, extent_mm):
        self.fdoc, self.view, self.rp, self.axis, self.at = fdoc, view, {}, {}, {}
        self.big = ft(extent_mm)
        self.failed = []
        # views where a lock can be made: the plan first, then the elevation seeing the planes edge-on
        # (elements above the plan cut plane)
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
        at = ft(at_mm)
        if axis == "x":
            a, b = XYZ(at, -self.big, 0), XYZ(at, self.big, 0)
        else:
            a, b = XYZ(-self.big, at, 0), XYZ(self.big, at, 0)
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

    def lock_faces(self, element, what, names=None):
        """Lock every vertical planar face of element to the reference plane it lies on. Returns the count.
        names: planes allowed for this element (default all); needed when two planes share a coordinate in
        the first type but must move apart in the others."""
        allowed = [n for n in self.rp if names is None or n in names]
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
                for name in allowed:
                    if self.axis[name] == axis and abs(self.at[name] - at) < 1e-4:
                        count += self.lock(name, face.Reference, what)
                        break
        return count

    def dimension(self, names, line_at_mm, label=None, equal=False):
        refs = ReferenceArray()
        for name in names:
            refs.Append(self.rp[name].GetReference())
        at = ft(line_at_mm)
        if self.axis[names[0]] == "x":
            line = Line.CreateBound(XYZ(-self.big, at, 0), XYZ(self.big, at, 0))
        else:
            line = Line.CreateBound(XYZ(at, -self.big, 0), XYZ(at, self.big, 0))
        dim = self.fdoc.FamilyCreate.NewDimension(self.view, line, refs)
        if label is not None:
            dim.FamilyLabel = label
        if equal:
            dim.AreSegmentsEqual = True
        return dim
