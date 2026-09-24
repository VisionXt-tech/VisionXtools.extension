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
assert T["UTI_TE_Funzione"] == ("I", ["OST_Areas"]) and "OST_Areas" in T["UTI_TE_Edificio"][1]
assert T["TA_TE_Titolo"] == ("I", ["OST_Sheets"]) and T["SAG_TE_Saggi"][1][0] == "OST_Walls"
assert "SGAS_View Status" not in T and "UTI_TE_Funzione" not in R.LEGACY_PARAMS
missing = [n for n, r in R.SHARED_PARAMS.items() if r["file"] in R.SHARED_FILES and n not in T]
assert not missing, missing  # every TXT parameter used by the ARC model has a binding
import os
assert all(os.path.exists(os.path.join(R.SHARED_DIR, f)) for f in R.SHARED_FILES.values())
assert R.wbs_l4_from_l6("i01") == "BG" and R.wbs_l4_from_l6("L02") == "AG"
assert R.room_abbreviation("L00A.BG.04", "Bagno") == "BG" and R.room_abbreviation("12", "Cucina") == "K"
for pairs in R.WBS_EXPECTED.values():
    for l7, l8 in pairs:
        assert R.wbs_level_errors(["BMA", "FP", "AUL", "AG", "E0C", "L01", l7, l8]) == [], (l7, l8)
assert R.wbs_expected("MUR", "OST_Walls", "CV1.01A") == [("3IN", "402")]
assert R.wbs_expected("POR", "OST_Doors", "PR2.01B") == [("3AR", "508")]
assert R.wbs_expected("OGG", "OST_Casework", "AI2.01A") == [("9FF", "B11")]
assert R.wbs_expected("MUR", "OST_Walls", "")[0] == ("3AR", "502")
assert R.wbs_expected("MUR", "OST_Walls", "PV2.01L", ["FIN.07"]) == [("3AR", "506")]  # cladding
assert R.wbs_expected("MUR", "OST_Walls", "PV2.01M", ["CTG.05"]) == [("3AR", "502")]
assert R.wbs_expected("RIN", "OST_StairsRailing", "RG2.01A")[0] == ("3AR", "515")
assert R.wbs_expected("FAC", "OST_Walls", "")[0] == ("3IN", "404")
P = ["BMA", "FP", "AUL", "AG", "E0C", "L01", "3AR", "502"]
assert R.wbs_merge(["CMC", "BMA", "AU", "E0C", "L00A", "EDI", "PIO", ""], P) == P  # old scheme replaced
assert R.wbs_merge(["BMA", "FP", "AUL", "AG", "E0C", "L01", "9FF", ""], P)[6:] == ["9FF", ""]  # 502 not under 9FF
assert R.wbs_merge(["BMA", "", "", "", "", "", "3IN", "402"], P) == P[:6] + ["3IN", "402"]  # valid kept
assert R.wbs_merge(["", "", "", "", "", "", "", "474"], P[:7] + [""])[7] == "474"  # never cleared
assert R.wbs_merge(["", "", "", "", ""], P, 5) == P[:5]
assert "OST_SpecialityEquipment" not in R.OGG_CATS and "OST_Planting" not in R.OGG_CATS  # PIR R06 OGG
assert not any(b in T["WBS_TE_WBS Livello 1"][1] for b in ("OST_SpecialityEquipment", "OST_Planting"))
assert "OST_FurnitureSystems" not in R.OGG_CATS and "OST_FurnitureSystems" not in R.CAT_CODE
assert T["VI_TE_Tema della vista"][1] == ["OST_Views", "OST_Schedules"] and T["VI_TE_Argomento della vista"][1] == ["OST_Views"]
assert R.wbs_expected("MUR", "OST_Walls", "OLD", ["INT.01"], "OLD_AR_MUR_CV1.01B_INT.01-ISO.06_880") == [("3IN", "402")]
print("bmc_rules: all checks passed")
