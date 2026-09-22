"""Center Room reference points and Room Tags based on the room shape"""

__title__ = "Center\nRoom Tags"
__author__ = "Luca Rosati"

import math

from pyrevit import revit, DB
from pyrevit import script

logger = script.get_logger()

GRID = 25  # fallback search resolution per axis


def room_loops(room, opts):
    """Boundary loops as lists of (x, y), curves tessellated."""
    loops = []
    for loop in room.GetBoundarySegments(opts) or []:
        pts = []
        for seg in loop:
            pts.extend((p.X, p.Y) for p in list(seg.GetCurve().Tessellate())[:-1])
        if len(pts) >= 3:
            loops.append(pts)
    return loops


def edges(pts):
    return zip(pts, pts[1:] + pts[:1])


def loop_area_centroid(pts):
    a = cx = cy = 0.0
    for (x0, y0), (x1, y1) in edges(pts):
        c = x0 * y1 - x1 * y0
        a += c
        cx += (x0 + x1) * c
        cy += (y0 + y1) * c
    if not a:
        return 0.0, 0.0, 0.0
    return a / 2.0, cx / (3.0 * a), cy / (3.0 * a)


def centroid(loops):
    """Area centroid; largest loop is the outer boundary, the others are holes."""
    parts = sorted((loop_area_centroid(p) for p in loops), key=lambda t: -abs(t[0]))
    A = X = Y = 0.0
    for i, (a, cx, cy) in enumerate(parts):
        a = abs(a) if i == 0 else -abs(a)
        A += a
        X += a * cx
        Y += a * cy
    return (X / A, Y / A) if A else None


def inside(x, y, loops):
    """Even-odd rule over all loops, so holes are excluded."""
    c = False
    for pts in loops:
        for (x0, y0), (x1, y1) in edges(pts):
            if (y0 > y) != (y1 > y) and x < x0 + (y - y0) * (x1 - x0) / (y1 - y0):
                c = not c
    return c


def clearance(x, y, loops):
    best = float("inf")
    for pts in loops:
        for (x0, y0), (x1, y1) in edges(pts):
            dx, dy = x1 - x0, y1 - y0
            ll = dx * dx + dy * dy
            t = max(0.0, min(1.0, ((x - x0) * dx + (y - y0) * dy) / ll)) if ll else 0.0
            best = min(best, math.hypot(x - x0 - t * dx, y - y0 - t * dy))
    return best


def shape_center(loops):
    """Centroid if it sits well inside the room, else the most interior grid point.

    ponytail: grid search is a coarse pole-of-inaccessibility (GRID^2 samples);
    switch to polylabel if tags on very thin/complex rooms land off-center.
    """
    c = centroid(loops)
    xs = [p[0] for pts in loops for p in pts]
    ys = [p[1] for pts in loops for p in pts]
    x0, y0 = min(xs), min(ys)
    sx, sy = (max(xs) - x0) / GRID, (max(ys) - y0) / GRID
    best, best_d = None, 0.0
    for i in range(GRID):
        for j in range(GRID):
            x, y = x0 + (i + 0.5) * sx, y0 + (j + 0.5) * sy
            if inside(x, y, loops):
                d = clearance(x, y, loops)
                if d > best_d:
                    best, best_d = (x, y), d
    if c and inside(c[0], c[1], loops) and clearance(c[0], c[1], loops) >= best_d / 2.0:
        return c
    return best or c


def mid_z(room):
    """Mid height of the room, used to center tags in section views."""
    bbox = room.get_BoundingBox(None)
    if bbox:
        return (bbox.Min.Z + bbox.Max.Z) / 2.0
    return room.Location.Point.Z + (room.UnboundedHeight or 0.0) / 2.0


doc = revit.doc
opts = DB.SpatialElementBoundaryOptions()
rooms = [
    x
    for x in revit.query.get_elements_by_categories([DB.BuiltInCategory.OST_Rooms])
    if x.Area > 0
]
tags = (
    DB.FilteredElementCollector(doc, doc.ActiveView.Id)
    .OfCategory(DB.BuiltInCategory.OST_RoomTags)
    .WhereElementIsNotElementType()
    .ToElements()
)

if rooms:
    try:
        with revit.Transaction("Center Rooms and Tags"):
            for room in rooms:
                try:
                    loops = room_loops(room, opts)
                    center = shape_center(loops) if loops else None
                    if not center:
                        continue
                    cur = room.Location.Point
                    room.Location.Move(DB.XYZ(center[0] - cur.X, center[1] - cur.Y, 0))
                except Exception as ex:
                    logger.warning("Room {}: {}".format(room.Id, ex))

            doc.Regenerate()

            view = doc.ActiveView
            right, up = view.RightDirection, view.UpDirection

            for tag in tags:
                room = tag.Room  # None for tags of linked rooms
                if not room or not room.Location or not tag.Location:
                    continue
                p = room.Location.Point
                target = DB.XYZ(p.X, p.Y, mid_z(room))
                delta = target - tag.Location.Point
                # keep the tag on the view plane: in a plan this is the XY move,
                # in a section it centers horizontally and vertically instead
                move = right.Multiply(delta.DotProduct(right)).Add(
                    up.Multiply(delta.DotProduct(up))
                )
                tag.Location.Move(move)
    except Exception as ex:
        logger.error("Center Room Tags failed: {}".format(ex))
