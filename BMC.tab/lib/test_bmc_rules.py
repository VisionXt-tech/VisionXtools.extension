"""Self-check of the pure logic in bmc_rules, runnable without Revit: python test_bmc_rules.py"""

import sys
import types


class _Any(object):
    """Stand-in for BuiltInCategory / BuiltInParameter: every attribute is its own name."""

    def __getattr__(self, name):
        return name


db = types.ModuleType("Autodesk.Revit.DB")
db.BuiltInCategory = _Any()
db.BuiltInParameter = _Any()
for name in ("Autodesk", "Autodesk.Revit"):
    sys.modules[name] = types.ModuleType(name)
sys.modules["Autodesk.Revit.DB"] = db

import bmc_rules as R  # noqa: E402

R.load_data()

ok = ["BMA", "FP", "AUL", "AG", "E0C", "L01", "3AR", "502"]
assert R.wbs_level_errors(ok) == []
assert (
    R.wbs_level_errors(["CMC", "BMA", "AU", "E0C", "B01A", "EDI", "PIV", ""])[0]
    == "L1 `CMC` non ammesso"
)
assert R.wbs_level_errors(ok[:7] + ["502,506"]) == []  # level 8 list, same L7
assert "appartiene a L7" in R.wbs_level_errors(ok[:7] + ["B10"])[0]
assert R.wbs_level_errors(["nd"] * 8) == []
assert R.wbs_l6_from_level("E0C_AR_L01A_pf_+0.59") == "L01"
assert R.wbs_l6_from_level("E0C_AR_B01A_pf_-8.33") == "i01"
assert R.wbs_l6_from_level("COO_LIVELLO ZERO") is None
assert R.RE_FILE.match("CMC-BMA-AU-XXX-PE-AR-MOD-M3-01")
assert R.RE_FILE.match("CMC-BMA-AU-XXX-PE-AR-MOD-M3-01_l.rosati")
assert R.RE_VIEW_TEMPLATE.match("PA_WIP_100_Architettonico")
assert R.RE_VIEW_SHEET.match("PA100-L01a.Pianta piano primo")
assert R.RE_VIEW_SCHEDULE.match("ABCOO-QU_Muri.QTO")
assert ("Stato di Progetto", "Comparativa") in R.VIEW_PHASING
assert len(R.all_pir_parameters()) > 40 and "Keynote" not in R.pir_names(
    R.all_pir_parameters()
)
assert all(name in R.SHARED_PARAMS for name, _ in R.all_pir_parameters()), [
    n for n, _ in R.all_pir_parameters() if n not in R.SHARED_PARAMS
]
T = R.binding_targets()
assert T["EPU_TE_Codice EPU"][0] == "T" and "OST_Walls" in T["EPU_TE_Codice EPU"][1]
assert (
    T["UTI_TE_Funzione"] == ("I", ["OST_Areas"])
    and "OST_Areas" in T["UTI_TE_Edificio"][1]
)
assert (
    T["TA_TE_Titolo"] == ("I", ["OST_Sheets"])
    and T["SAG_TE_Saggi"][1][0] == "OST_Walls"
)
assert "SGAS_View Status" not in T and "UTI_TE_Funzione" not in R.LEGACY_PARAMS
missing = [
    n for n, r in R.SHARED_PARAMS.items() if r["file"] in R.SHARED_FILES and n not in T
]
assert not missing, missing  # every TXT parameter used by the ARC model has a binding
import os

assert all(
    os.path.exists(os.path.join(R.SHARED_DIR, f)) for f in R.SHARED_FILES.values()
)
assert R.wbs_l4_from_l6("i01") == "BG" and R.wbs_l4_from_l6("L02") == "AG"
assert (
    R.room_abbreviation("L00A.BG.04", "Bagno") == "BG"
    and R.room_abbreviation("12", "Cucina") == "K"
)
for pairs in R.WBS_EXPECTED.values():
    for l7, l8 in pairs:
        assert (
            R.wbs_level_errors(["BMA", "FP", "AUL", "AG", "E0C", "L01", l7, l8]) == []
        ), (l7, l8)
assert R.wbs_expected("MUR", "OST_Walls", "CV1.01A") == [("3IN", "402")]
assert R.wbs_expected("POR", "OST_Doors", "PR2.01B") == [("3AR", "508")]
assert R.wbs_expected("OGG", "OST_Casework", "AI2.01A") == [("9FF", "B11")]
assert R.wbs_expected("MUR", "OST_Walls", "")[0] == ("3AR", "502")
assert R.wbs_expected("MUR", "OST_Walls", "PV2.01L", ["FIN.07"]) == [
    ("3AR", "506")
]  # cladding
assert R.wbs_expected("MUR", "OST_Walls", "PV2.01M", ["CTG.05"]) == [("3AR", "502")]
assert R.wbs_expected("RIN", "OST_StairsRailing", "RG2.01A") == [
    ("3AR", "515")
]  # validated 25/09
assert R.wbs_expected("OGG", "OST_Stairs", "PI2.01A") == [("3AR", "506")]
assert R.wbs_expected("POR", "OST_Doors", "PR1.01A") == [
    ("3IN", "402")
]  # external door: opaque facade
assert R.wbs_expected("POR", "OST_Doors", "PR1.01A", in_curtain_wall=True) == [
    ("3IN", "404")
]
assert R.wbs_expected("POR", "OST_Doors", "PR2.01B", in_curtain_wall=True) == [
    ("3IN", "404")
]
assert R.wbs_expected("PAV", "OST_Floors", "CS1.01A") == [("3IN", "454")]
assert (
    R.strat_cat_code("OST_Floors", "CS") == "COP"
    and R.strat_cat_code("OST_Floors", "PO") == "PAV"
)
assert R.strat_cat_code("OST_Walls", "CS") == "MUR"
assert R.unique_mark("P", 7) == "P07" and R.unique_mark("P", 118) == "P118"
assert (
    R.RE_UNIQUE.match("P07")
    and R.RE_UNIQUE.match("F118")
    and not R.RE_UNIQUE.match("P7")
)
assert R.host_thickness_code("PV2.01K", 111.99999) == "PV2.01K_112"
assert R.host_thickness_code("PV2.01L", 12.0) == "PV2.01L_012"
assert (
    R.host_thickness_code("CV1.01A", 203.4999999) == "CV1.01A_204"
)  # half up after 0.001 mm
assert R.wbs_expected("FAC", "OST_Walls", "")[0] == ("3IN", "404")
P = ["BMA", "FP", "AUL", "AG", "E0C", "L01", "3AR", "502"]
assert (
    R.wbs_merge(["CMC", "BMA", "AU", "E0C", "L00A", "EDI", "PIO", ""], P) == P
)  # old scheme replaced
assert R.wbs_merge(["BMA", "FP", "AUL", "AG", "E0C", "L01", "9FF", ""], P)[6:] == [
    "9FF",
    "",
]  # 502 not under 9FF
assert R.wbs_merge(["BMA", "", "", "", "", "", "3IN", "402"], P) == P[:6] + [
    "3IN",
    "402",
]  # valid kept
assert (
    R.wbs_merge(["", "", "", "", "", "", "", "474"], P[:7] + [""])[7] == "474"
)  # never cleared
assert R.wbs_merge(["", "", "", "", ""], P, 5) == P[:5]
assert (
    "OST_SpecialityEquipment" not in R.OGG_CATS and "OST_Planting" not in R.OGG_CATS
)  # PIR R06 OGG
assert not any(
    b in T["WBS_TE_WBS Livello 1"][1]
    for b in ("OST_SpecialityEquipment", "OST_Planting")
)
assert (
    "OST_FurnitureSystems" not in R.OGG_CATS
    and "OST_FurnitureSystems" not in R.CAT_CODE
)
assert T["VI_TE_Tema della vista"][1] == ["OST_Views", "OST_Schedules"] and T[
    "VI_TE_Argomento della vista"
][1] == ["OST_Views"]
assert R.wbs_expected(
    "MUR", "OST_Walls", "OLD", ["INT.01"], "OLD_AR_MUR_CV1.01B_INT.01-ISO.06_880"
) == [("3IN", "402")]
# clockwise from north around the centroid: N, E, S, W whatever the input order
assert R.clockwise_order([(0, -5, "S"), (-5, 0, "W"), (0, 5, "N"), (5, 0, "E")]) == [
    "N",
    "E",
    "S",
    "W",
]
assert R.clockwise_order([(-1, 5, "NW"), (1, 5, "NE"), (0, -5, "S")]) == [
    "NE",
    "S",
    "NW",
]
assert R.RE_HOST_THK.match("PV2.01K_112") and not R.RE_HOST_THK.match("OLD_160")
assert [R.layer_thickness_cm(w) for w in (15, 12.5, 450, 6, 75)] == [
    "1,5",
    "1,25",
    "45",
    "0,6",
    "7,5",
]
_layers = [
    ("GAS.01_Intercapedine d'aria", 10),
    ("ISO.01_Pannelli in lana di roccia, cond. 0,035 W/mk", 100),
    ("BAR.02_Barriera al vapore", 0),
    ("TIN.01_Tinteggiatura", 0.0001),
]
assert R.stratigraphy_description(
    "Controparete verso vano ascensore.  Composizione: GAS.01_Intercapedine d'aria (1 cm) + ISO.01_Pannelli ... (10 cm)",
    _layers,
) == (
    "Controparete verso vano ascensore. Composizione: GAS.01_Intercapedine d'aria (1 cm) + "
    "ISO.01_Pannelli in lana di roccia, cond. 0,035 W/mk (10 cm) + BAR.02_Barriera al vapore + TIN.01_Tinteggiatura"
)
assert (
    R.stratigraphy_description("", _layers[:1])
    == "Composizione: GAS.01_Intercapedine d'aria (1 cm)"
)
assert (
    R.stratigraphy_description("nd", _layers[:1])
    == "Composizione: GAS.01_Intercapedine d'aria (1 cm)"
)
assert R.stratigraphy_description("Muro esistente.", _layers[:1]).startswith(
    "Muro esistente. Composizione:"
)
# lining along x (4 m, 111 mm), room on +y: structure searched on -y (side -1); structural wall 300 mm
_A = ((0, 0), (4000, 0), 55.5)
_gap240 = -(55.5 + 240 + 150)
_r = R.wall_behind(
    _A[0], _A[1], _A[2], -1, (-500, _gap240), (3000, _gap240), 150, 500, 300
)
assert (
    abs(_r[0] - 240) < 1e-6 and _r[1] == 3000
)  # gap 240, overlap clipped to the lining start
_r = R.wall_behind(_A[0], _A[1], _A[2], -1, (4000, -205.5), (0, -205.5), 150, 500, 300)
assert _r == (0.0, 4000)  # against the structure, reversed direction
assert (
    R.wall_behind(_A[0], _A[1], _A[2], -1, (0, 205.5), (4000, 205.5), 150, 500, 300)
    is None
)  # room side
assert R.wall_behind(
    _A[0], _A[1], _A[2], 0, (0, 205.5), (4000, 205.5), 150, 500, 300
) == (0.0, 4000)
assert (
    R.wall_behind(_A[0], _A[1], _A[2], -1, (0, -300), (0, -3000), 150, 500, 300) is None
)  # perpendicular
assert (
    R.wall_behind(_A[0], _A[1], _A[2], -1, (0, -900), (4000, -900), 150, 500, 300)
    is None
)  # too far
assert (
    R.wall_behind(_A[0], _A[1], _A[2], -1, (3900, -300), (6000, -300), 150, 500, 300)
    is None
)  # overlap 100
assert R.vvf_room_kinds("L02A.VS.17", "Filtro fumo") == ("VS", "FF")
assert R.vvf_room_kinds("L00A.VS.17", "Vano Scala") == ("VS", "VS")
assert R.vvf_room_kinds("L00A.CZ.11", "Circolazione") == ("", "")
print("bmc_rules: all checks passed")
