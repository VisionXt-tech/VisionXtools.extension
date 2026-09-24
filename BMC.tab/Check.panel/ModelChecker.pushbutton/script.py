"""Check the BMC architectural model (Progetto Esecutivo) against PGI R06, PIR R06 and WBS R06 and report errors and warnings"""

__title__ = "Model\nChecker"
__author__ = "Luca Rosati"
__guida__ = "Controlla il modello rispetto a PGI, PIR e WBS R06 (nomi, parametri, WBS, workset, livelli, fasi, viste) e salva il report CSV in Documenti. Non modifica nulla."

# Read-only: no Transaction is opened, the model is never modified.
# Engine: IronPython 2.7 (pyRevit default). ASCII-only source.
# Rules: BMC.tab/lib/bmc_rules.py (PGI R06 22/09/2026). Analysis: docs/_corrente/R06_analisi.md.
# Output: pyRevit report with element links + CSV in Documents (evidence of the LV1 check, PGI 3.14.2.1).

import clr
import codecs
import datetime
import os

clr.AddReference("System.Core")
from System.Collections.Generic import HashSet

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

from pyrevit import script

import bmc_rules as R
from bmc_utils import (
    ename,
    idv,
    instances,
    layer_codes,
    level_name,
    param_state,
    param_state_any,
    to_m,
    to_mm,
    type_mark,
    used_types,
    workset_name,
)

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
output = script.get_output()

OK, WARN, ERR, FAIL = 0, 1, 2, 3
LABEL = {OK: "OK", WARN: "AVVISO", ERR: "ERRORE", FAIL: "ERRORE SCRIPT"}
MAX_ROWS = 50
MAX_SELECT = 1000

# ---------------------------------------------------------------- check registry

checks = []
runners = []


class Check(object):
    def __init__(self, cid, title, source):
        self.cid, self.title, self.source = cid, title, source
        self.status = OK
        self.items = []  # (element id | [ids] | None, text)
        self.note = None

    def add(self, text, element_id=None, severity=ERR):
        self.items.append((element_id, text))
        self.status = max(self.status, severity)


def check(cid, title, source):
    def wrap(func):
        def runner():
            c = Check(cid, title, source)
            try:
                func(c)
            except Exception as ex:
                c.status = FAIL
                c.note = "Errore nello script: %s" % ex
            checks.append(c)

        runners.append(runner)
        return func

    return wrap


def room_label(room):
    return "Locale `%s %s`" % (
        room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString(),
        room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString(),
    )


# ---------------------------------------------------------------- 00 model


@check("00.1", "Versione Revit", "PGI 3.1.2")
def c_version(c):
    if app.VersionNumber != "2024":
        c.add("Versione Revit %s (attesa 2024)" % app.VersionNumber)


@check("00.2", "Nome file modello", "PGI 3.11.3")
def c_file_name(c):
    if not R.RE_FILE.match(doc.Title):
        c.add(
            "Nome file `%s` non conforme a FONDO-COMMESSA-DEST-EDIFICIO-FASE-DISC-MOD-SOTTOTIPO-NN"
            % doc.Title
        )


@check("00.3", "Dimensione del file", "PGI 3.11.1 (max %d MB)" % R.MAX_FILE_MB)
def c_file_size(c):
    path = doc.PathName
    if not path or not os.path.exists(path):
        c.note = "File non salvato in locale: dimensione non verificabile"
        return
    mb = os.path.getsize(path) / 1048576.0
    if mb > R.MAX_FILE_MB:
        c.add("%.0f MB (massimo %d MB)" % (mb, R.MAX_FILE_MB))
    else:
        c.note = "%.0f MB" % mb


@check("00.4", "Parametri condivisi: PIR R06 vs TXT R06", "PIR R06; PGI 3.12.1")
def c_params(c):
    bound = {}
    it = doc.ParameterBindings.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        guid = None
        spe = doc.GetElement(it.Key.Id)
        if isinstance(spe, SharedParameterElement):
            guid = str(spe.GuidValue).lower()
        is_type = isinstance(it.Current, TypeBinding)
        bound.setdefault(it.Key.Name, []).append((guid, is_type))
    for name, level in R.all_pir_parameters():
        ref = R.SHARED_PARAMS.get(name)
        if name not in bound:
            c.add(
                "`%s` mancante (TXT %s): Parametri > Riemetti parametri R06"
                % (name, ref["file"] if ref else "?")
            )
            continue
        for guid, is_type in bound[name]:
            if ref and guid != ref["guid"].lower():
                c.add(
                    "`%s`: GUID diverso dal TXT %s (parametro da sostituire)"
                    % (name, ref["file"])
                )
            if is_type != (level == "T"):
                c.add(
                    "`%s`: associato come %s, il PIR lo vuole di %s"
                    % (
                        name,
                        "tipo" if is_type else "istanza",
                        "tipo" if level == "T" else "istanza",
                    ),
                    severity=WARN,
                )
    for name, entries in sorted(bound.items()):
        if len(entries) > 1:
            c.add("`%s` presente %d volte" % (name, len(entries)))
        if name in R.LEGACY_PARAMS:
            c.add("`%s` non previsto dal PIR R06: da eliminare" % name, severity=WARN)


# ---------------------------------------------------------------- 01 coordinates


@check("01.1", "Project Base Point", "PGI 3.3")
def c_pbp(c):
    bp = BasePoint.GetProjectBasePoint(doc)
    values = {
        "north_south": to_m(
            bp.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble()
        ),
        "east_west": to_m(
            bp.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble()
        ),
        "elevation": to_m(
            bp.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble()
        ),
        "angle": bp.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsDouble()
        * 180.0
        / 3.141592653589793,
    }
    for key, expected in R.PBP_EXPECTED.items():
        if abs(values[key] - expected) > R.COORD_TOL:
            c.add("%s = %.3f (atteso %.3f)" % (key, values[key], expected), bp.Id)


@check("01.2", "Link RVT posizionati da coordinate condivise", "PGI 3.3")
def c_links_shared(c):
    for li in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        if "<Not Shared>" in ename(li):
            c.add(
                "Link `%s` non posizionato da coordinate condivise" % ename(li), li.Id
            )


# ---------------------------------------------------------------- 02 phases


@check("02.1", "Fasi di progetto", "PGI 3.12.4")
def c_phases(c):
    names = [p.Name for p in doc.Phases]
    if names != R.PHASES_EXPECTED:
        c.add(
            "Fasi nel modello: %s - attese: %s"
            % (", ".join(names), ", ".join(R.PHASES_EXPECTED))
        )


@check("02.2", "Filtri fase", "PGI 3.12.5")
def c_phase_filters(c):
    statuses = [
        ElementOnPhaseStatus.New,
        ElementOnPhaseStatus.Existing,
        ElementOnPhaseStatus.Demolished,
        ElementOnPhaseStatus.Temporary,
    ]
    filters = dict(
        (pf.Name, pf) for pf in FilteredElementCollector(doc).OfClass(PhaseFilter)
    )
    for aliases, expected in R.PHASE_FILTERS_EXPECTED:
        found = [filters[n] for n in aliases if n in filters]
        if not found:
            c.add("Filtro fase `%s` mancante" % aliases[0])
            continue
        actual = tuple(str(found[0].GetPhaseStatusPresentation(s)) for s in statuses)
        if actual != expected:
            c.add(
                "Filtro `%s`: New/Existing/Demolished/Temporary = %s (atteso %s)"
                % (found[0].Name, "/".join(actual), "/".join(expected)),
                found[0].Id,
            )


@check("02.3", "Nessun elemento creato in Strip-out / Post Strip-out", "PGI 3.12.4")
def c_phase_created(c):
    bad = [idv(p.Id) for p in doc.Phases if p.Name in R.PHASES_NO_CREATION]
    if not bad:
        c.note = "Fasi Strip-out non presenti nel modello: controllo non applicabile"
        return
    ids = []
    for el in FilteredElementCollector(doc).WhereElementIsNotElementType():
        p = el.get_Parameter(BuiltInParameter.PHASE_CREATED)
        if p is not None and idv(p.AsElementId()) in bad:
            ids.append(el.Id)
    if ids:
        c.add("%d elementi creati in una fase di strip-out" % len(ids), ids)


@check("02.4", "Combinazione fase / filtro fase nelle viste", "PGI 3.12.6")
def c_view_phasing(c):
    wrong = {}
    for v in real_views():
        pp = v.get_Parameter(BuiltInParameter.VIEW_PHASE)
        pf = v.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER)
        if pp is None or pf is None or pp.AsElementId() == ElementId.InvalidElementId:
            continue
        phase = doc.GetElement(pp.AsElementId()).Name
        pf_el = doc.GetElement(pf.AsElementId())
        pair = (phase, pf_el.Name if pf_el is not None else None)
        if pair not in R.VIEW_PHASING:
            wrong.setdefault(pair, []).append(v.Id)
    for (phase, pfilter), ids in sorted(wrong.items()):
        c.add(
            "Fase `%s` con filtro `%s` non prevista (%d viste)"
            % (phase, pfilter or "None", len(ids)),
            ids,
            WARN,
        )


# ---------------------------------------------------------------- 03 worksets


@check(
    "03.1",
    "Workset: coordinamento `_...`, %s, %s" % (R.WS_MODEL, R.WS_AREAS),
    "PGI 3.11.8; decisione 23/09",
)
def c_worksets(c):
    names = [
        w.Name for w in FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)
    ]
    if R.WS_MODEL not in names:
        c.add(
            "Workset di modellazione `%s` mancante (Fix > Workset unico)" % R.WS_MODEL
        )
    for n in names:
        if n in R.WS_COORD or n in (R.WS_MODEL, R.WS_AREAS):
            continue
        if n.strip() in R.WS_COORD:
            c.add("Workset `%s`: spazi iniziali/finali nel nome" % n, severity=WARN)
        else:
            c.add(
                "Workset `%s`: elementi da unire in `%s` (Fix > Workset unico)"
                % (n, R.WS_MODEL)
            )
    for n in R.WS_COORD:
        if n not in [x.strip() for x in names]:
            c.add("Workset di coordinamento `%s` mancante" % n, severity=WARN)


@check("03.2", "Link, griglie, livelli e aree nei workset dedicati", "PGI 3.11.8")
def c_ws_placement(c):
    for li in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        expected = R.WS_IFC if ".ifc" in ename(li).lower() else R.WS_RVT
        if workset_name(doc, li) != expected:
            c.add(
                "Link `%s` su `%s` (atteso `%s`)"
                % (ename(li), workset_name(doc, li), expected),
                li.Id,
            )
    for ii in FilteredElementCollector(doc).OfClass(ImportInstance):
        if ii.IsLinked and workset_name(doc, ii) != R.WS_DWG:
            c.add(
                "DWG collegato su `%s` (atteso `%s`)"
                % (workset_name(doc, ii), R.WS_DWG),
                ii.Id,
            )
    for datum in list(FilteredElementCollector(doc).OfClass(Level)) + list(
        FilteredElementCollector(doc).OfClass(Grid)
    ):
        if workset_name(doc, datum) != R.WS_GRIDS:
            c.add(
                "`%s` su `%s` (atteso `%s`)"
                % (datum.Name, workset_name(doc, datum), R.WS_GRIDS),
                datum.Id,
            )
    areas = [
        a.Id
        for a in instances(doc, [BuiltInCategory.OST_Areas])
        if workset_name(doc, a) != R.WS_AREAS
    ]
    if areas:
        c.add(
            "%d aree non su `%s` (Fix > Workset unico)" % (len(areas), R.WS_AREAS),
            areas,
        )


# ---------------------------------------------------------------- 04 levels


@check("04.1", "Nome livelli e quota nel nome", "PGI 3.11.4; riunione 21/09")
def c_levels(c):
    for lv in FilteredElementCollector(doc).OfClass(Level):
        if lv.Name == R.LEVEL_SEA or R.RE_LEVEL_COO.match(lv.Name):
            continue
        if lv.Name.startswith("COO"):
            c.add("Livello di coordinamento `%s`: rinominare COO_..." % lv.Name, lv.Id)
            continue
        m = R.RE_LEVEL.match(lv.Name)
        if not m:
            c.add(
                "Livello `%s` non conforme a EDIFICIO_AR_CODICE_pf/pr/int_+Q.QQ"
                % lv.Name,
                lv.Id,
            )
            continue
        elevation = to_m(lv.Elevation)
        if abs(float(m.group(4)) - elevation) > 0.005:
            c.add(
                "Livello `%s`: quota nel nome %s, reale %.3f"
                % (lv.Name, m.group(4), elevation),
                lv.Id,
            )


@check("04.2", "Livelli e griglie monitorati e bloccati", "PGI 3.11.4")
def c_datums_monitor(c):
    for datum in list(FilteredElementCollector(doc).OfClass(Level)) + list(
        FilteredElementCollector(doc).OfClass(Grid)
    ):
        if not datum.IsMonitoringLinkElement():
            c.add("`%s` non monitorato (Copy/Monitor)" % datum.Name, datum.Id)
        if not datum.Pinned:
            c.add("`%s` non bloccato (pin)" % datum.Name, datum.Id, WARN)


@check("04.3", "Livelli e griglie con Scope Box", "PGI 3.11.1")
def c_scope_box(c):
    for datum in list(FilteredElementCollector(doc).OfClass(Level)) + list(
        FilteredElementCollector(doc).OfClass(Grid)
    ):
        p = datum.get_Parameter(BuiltInParameter.DATUM_VOLUME_OF_INTEREST)
        if p is None or p.AsElementId() == ElementId.InvalidElementId:
            c.add("`%s` senza Scope Box" % datum.Name, datum.Id, WARN)


@check("04.4", "Livelli pr/int con flag Strutturale", "PGI 3.11.4")
def c_level_structural(c):
    for lv in FilteredElementCollector(doc).OfClass(Level):
        m = R.RE_LEVEL.match(lv.Name)
        if not m:
            continue
        p = lv.get_Parameter(BuiltInParameter.LEVEL_IS_STRUCTURAL)
        is_structural = p is not None and p.AsInteger() == 1
        if m.group(3) in R.LEVEL_STRUCTURAL_KINDS and not is_structural:
            c.add("`%s`: livello strutturale senza flag Strutturale" % lv.Name, lv.Id)
        elif m.group(3) == "pf" and is_structural:
            c.add("`%s`: piano finito con flag Strutturale" % lv.Name, lv.Id, WARN)


# ---------------------------------------------------------------- 05 naming


def code_of(match):
    return "%s%s.%s%s" % (
        match.group(2),
        match.group(3),
        match.group(4),
        match.group(5),
    )


@check("05.1", "Nome tipi stratigrafici e spessore", "PGI 3.11.5.3, 3.11.7")
def c_strat_names(c):
    for t, n in used_types(doc, R.STRAT_CATS):
        if not isinstance(t, HostObjAttributes):
            continue
        if isinstance(t, WallType) and t.Kind != WallKind.Basic:
            continue
        cs = t.GetCompoundStructure()
        if cs is None:
            continue
        name = ename(t)
        if name.startswith("OLD_"):
            c.add(
                "Tipo obsoleto `%s` usato da %d elementi (sostituzione manuale)"
                % (name, n),
                t.Id,
            )
            continue
        m = R.RE_STRAT_TYPE.match(name)
        if not m:
            c.add(
                "`%s` non conforme a AR_CAT_AAX.XXA_Composizione_Spessore (%d elementi)"
                % (name, n),
                t.Id,
            )
            continue
        expected = R.STRAT_CAT_CODE.get(t.Category.BuiltInCategory)
        if expected and m.group(1) != expected:
            c.add(
                "`%s`: codice categoria %s (atteso %s)" % (name, m.group(1), expected),
                t.Id,
            )
        if m.group(2) not in R.STRAT_CODES:
            c.add(
                "`%s`: codice stratigrafia %s non in tabella PGI 3.11.7"
                % (name, m.group(2)),
                t.Id,
            )
        width = round(to_mm(cs.GetWidth()), 1)
        if abs(float(m.group(7)) - width) > 0.5:
            c.add(
                "`%s`: spessore nel nome %s mm, reale %s mm (%d elementi)"
                % (name, m.group(7), width, n),
                t.Id,
            )
        if type_mark(t) != code_of(m):
            c.add(
                "`%s`: Type Mark `%s` diverso dal codice nel nome"
                % (name, type_mark(t)),
                t.Id,
            )


@check(
    "05.2",
    "Nome famiglie e tipi caricabili, progressivo nel range",
    "PGI 3.11.5.1-2, 3.11.6",
)
def c_loadable_names(c):
    seen = set()
    for t, n in used_types(
        doc,
        R.LOADABLE_CATS + [BuiltInCategory.OST_Railings, BuiltInCategory.OST_Stairs],
    ):
        name = ename(t)
        expected = R.CAT_CODE.get(t.Category.BuiltInCategory) if t.Category else None
        fm = None
        if isinstance(t, FamilySymbol):
            fam = t.Family
            if fam.IsInPlace:
                continue
            fm = R.RE_FAMILY.match(fam.Name)
            if idv(fam.Id) not in seen:
                seen.add(idv(fam.Id))
                if not fm:
                    c.add(
                        "Famiglia `%s` non conforme a AR_CAT_AAX.XXA_Descrizione"
                        % fam.Name,
                        fam.Id,
                    )
                elif int(fm.group(4)) % 10 != 0:
                    c.add(
                        "Famiglia `%s`: caposaldo %s non multiplo di 10"
                        % (fam.Name, fm.group(4)),
                        fam.Id,
                    )
        m = R.RE_TYPE.match(name)
        if not m:
            c.add(
                "Tipo `%s` non conforme a AR_CAT_AAX.XXA_Descrizione_Dimensioni (%d elementi)"
                % (name, n),
                t.Id,
            )
            continue
        if expected and m.group(1) != expected:
            c.add(
                "Tipo `%s`: codice categoria %s (atteso %s)"
                % (name, m.group(1), expected),
                t.Id,
            )
        if type_mark(t) != code_of(m):
            c.add(
                "Tipo `%s`: Type Mark `%s` diverso dal codice nel nome"
                % (name, type_mark(t)),
                t.Id,
            )
        if (
            fm
            and int(fm.group(4)) % 10 == 0
            and int(m.group(4)) // 10 != int(fm.group(4)) // 10
        ):
            c.add(
                "Tipo `%s`: progressivo %s fuori dal range della famiglia (%s..%s)"
                % (name, m.group(4), fm.group(4), int(fm.group(4)) + 9),
                t.Id,
            )


@check("05.3", "Dimensioni nel nome di porte e finestre", "PGI 3.11.5.2")
def c_opening_dims(c):
    pairs = [
        (
            BuiltInCategory.OST_Doors,
            BuiltInParameter.DOOR_WIDTH,
            BuiltInParameter.DOOR_HEIGHT,
        ),
        (
            BuiltInCategory.OST_Windows,
            BuiltInParameter.WINDOW_WIDTH,
            BuiltInParameter.WINDOW_HEIGHT,
        ),
    ]
    for bic, bw, bh in pairs:
        for t, n in used_types(doc, [bic]):
            m = R.RE_TYPE.match(ename(t))
            if not m or "x" not in m.group(7):
                continue
            pw, ph = t.get_Parameter(bw), t.get_Parameter(bh)
            if pw is None or ph is None or not pw.HasValue or not ph.HasValue:
                continue
            nw, nh = [float(x) for x in m.group(7).split("x")[:2]]
            rw, rh = round(to_mm(pw.AsDouble())), round(to_mm(ph.AsDouble()))
            if abs(nw - rw) > 0.5 or abs(nh - rh) > 0.5:
                c.add(
                    "`%s`: nome %dx%d, reale %dx%d (%d elementi)"
                    % (ename(t), nw, nh, rw, rh, n),
                    t.Id,
                )


@check("05.4", "Type Mark univoco per tipo", "PGI 3.11.5")
def c_type_mark_unique(c):
    marks = {}
    for t in FilteredElementCollector(doc).WhereElementIsElementType():
        if t.Category is None or t.Category.CategoryType != CategoryType.Model:
            continue
        value = type_mark(t)
        if value:
            marks.setdefault((t.Category.Name, value), []).append(t)
    for (cat, value), types in sorted(marks.items()):
        if len(types) > 1:
            c.add(
                "Type Mark `%s` (%s) usato da %d tipi: %s"
                % (value, cat, len(types), ", ".join("`%s`" % ename(t) for t in types)),
                types[0].Id,
            )


@check(
    "05.5",
    "Keynote = codice nel nome del tipo = Type Mark",
    "riunione 21/09 (verifica pre-consegna)",
)
def c_keynote(c):
    for t in FilteredElementCollector(doc).WhereElementIsElementType():
        if t.Category is None or t.Category.CategoryType != CategoryType.Model:
            continue
        m = R.RE_CODE_PREFIX.match(ename(t))
        if not m:
            continue
        kn = t.get_Parameter(BuiltInParameter.KEYNOTE_PARAM)
        keynote = (kn.AsString() or "").strip() if kn is not None else ""
        if keynote != m.group(1):
            c.add(
                "`%s`: Keynote `%s` (attesa `%s`)" % (ename(t), keynote, m.group(1)),
                t.Id,
            )
        if type_mark(t) != m.group(2):
            c.add(
                "`%s`: Type Mark `%s` (atteso `%s`)"
                % (ename(t), type_mark(t), m.group(2)),
                t.Id,
            )


# ---------------------------------------------------------------- 06 PIR parameters


def group_elements(bics, wall_kind):
    elements = instances(doc, bics)
    if wall_kind:
        want = WallKind.Curtain if wall_kind == "curtain" else WallKind.Basic
        elements = [
            e for e in elements if isinstance(e, Wall) and e.WallType.Kind == want
        ]
    return elements


def pir_check(c, bics, wall_kind, params):
    elements = group_elements(bics, wall_kind)
    missing, failures = set(), {}
    for el in elements:
        for name, level in params:
            target, state, _ = param_state_any(doc, el, name, level)
            if state == "missing":
                missing.add(name)
            elif state == "empty":
                failures.setdefault(name, {})[idv(target.Id)] = target.Id
    for name in sorted(missing):
        c.add("`%s`: parametro assente (vedi 00.4)" % name)
    for name in sorted(failures):
        ids = list(failures[name].values())
        c.add("`%s`: non compilato su %d elementi/tipi" % (name, len(ids)), ids)
    if not elements:
        c.note = "Nessun elemento nel modello"


for _code, _label, _bics, _kind, _params in R.PIR_GROUPS:

    def _make(bics=_bics, kind=_kind, params=_params):
        return lambda c: pir_check(c, bics, kind, params)

    check(
        "06.%s" % _code,
        "Compilazione PIR - %s %s" % (_code, _label),
        "PIR R06 P. Esecutivo",
    )(_make())


def wbs_elements():
    result = []
    for code, _label, bics, kind, params in R.PIR_GROUPS:
        if R.WBS_LEVELS[0] in R.pir_names(params):
            result.extend(group_elements(bics, kind))
    return result


def wbs_values(el):
    values = []
    for name in R.WBS_LEVELS:
        state, value = param_state(el, name)
        values.append(value if state in ("value", "nd") else "")
    return values


@check(
    "06.WBS",
    "WBS: valori ammessi, L8 figlio di L7, concatenato",
    "WBS R06; PGI 3.14.2.2 LV2.d",
)
def c_wbs(c):
    out_of_domain, concat_wrong = {}, []
    for el in wbs_elements():
        values = wbs_values(el)
        for msg in R.wbs_level_errors(values):
            out_of_domain.setdefault(msg, []).append(el.Id)
        state, concat = param_state(el, R.WBS_CONCAT)
        filled = [v for v in values if v]
        if state == "value" and filled and concat != R.WBS_SEP.join(filled):
            concat_wrong.append(el.Id)
    for msg, ids in sorted(out_of_domain.items(), key=lambda kv: -len(kv[1])):
        c.add("%s (%d elementi)" % (msg, len(ids)), ids)
    if concat_wrong:
        c.add(
            "%d elementi con WBS_TE_WBS diverso dalla concatenazione dei livelli"
            % len(concat_wrong),
            concat_wrong,
            WARN,
        )


@check(
    "06.WBS8",
    "WBS: L7/L8 coerenti con il tipo di elemento",
    "WBS R06; regola bmc_rules.WBS_EXPECTED (proposta da validare)",
)
def c_wbs_kind(c):
    wrong = {}
    for code, _label, bics, kind, params in R.PIR_GROUPS:
        if R.WBS_LEVELS[7] not in R.pir_names(params):
            continue
        for el in group_elements(bics, kind):
            values = wbs_values(el)
            if not values[6] or R.wbs_level_errors(values):
                continue  # empty or not admissible: reported by 06.xxx and 06.WBS
            el_type = doc.GetElement(el.GetTypeId())
            expected = R.wbs_expected(
                code,
                el.Category.BuiltInCategory,
                type_mark(el_type) if el_type else "",
                layer_codes(doc, el_type) if el_type else [],
                ename(el_type) if el_type else "",
            )
            if expected and (values[6], values[7]) not in expected:
                key = (
                    ename(el_type) if el_type else "-",
                    "%s/%s" % (values[6], values[7]),
                    ", ".join("%s/%s" % p for p in expected),
                )
                wrong.setdefault(key, []).append(el.Id)
    for (type_name, pair, expected), ids in sorted(wrong.items()):
        c.add(
            "`%s`: L7/L8 %s, atteso %s (%d elementi)"
            % (type_name, pair, expected, len(ids)),
            ids,
            WARN,
        )


@check("06.ND", "Valori nd e valori numerici", "riunione 21/09 (regole nd)")
def c_nd(c):
    spelling, wbs_mixed, zeros = {}, [], {}
    for _code, _label, bics, kind, params in R.PIR_GROUPS:
        for el in group_elements(bics, kind):
            wbs = []
            for name, level in params:
                _, state, value = param_state_any(doc, el, name, level)
                if state not in ("value", "nd"):
                    continue
                text = value.strip() if isinstance(value, basestring) else value
                if text in R.ND_VARIANTS:
                    spelling.setdefault(name, []).append(el.Id)
                if name in R.WBS_LEVELS:
                    wbs.append(text)
                if "_NU_" in name and text in (0, "0", "0.00", "0,00"):
                    zeros.setdefault(name, []).append(el.Id)
            if R.ND in wbs and any(v != R.ND for v in wbs):
                wbs_mixed.append(el.Id)
    for name, ids in sorted(spelling.items()):
        c.add(
            "`%s`: grafia non standard su %d elementi (usare `%s`)"
            % (name, len(ids), R.ND),
            ids,
            WARN,
        )
    if wbs_mixed:
        c.add(
            "%d elementi con WBS in parte nd: se un campo WBS e' nd lo sono tutti"
            % len(wbs_mixed),
            wbs_mixed,
        )
    for name, ids in sorted(zeros.items()):
        c.add(
            "`%s` = 0 su %d elementi: ammesso solo se giustificato" % (name, len(ids)),
            ids,
            WARN,
        )


@check("06.HOST", "Codice elemento tecnico HOST = Type Mark dell'host", "PIR R06 SPE")
def c_host_code(c):
    for el in instances(
        doc,
        [
            BuiltInCategory.OST_Doors,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_CurtainWallPanels,
        ],
    ):
        host = getattr(el, "Host", None)
        if host is None:
            continue
        state, value = param_state(el, R.P_HOST)
        if state != "value":
            continue  # empty values are reported by 06.POR / 06.FIN / 06.PFC
        expected = type_mark(doc.GetElement(host.GetTypeId()))
        if value != expected:
            c.add("Valore `%s`, Type Mark host `%s`" % (value, expected), el.Id)


# ---------------------------------------------------------------- 07 rooms


@check("07.1", "Locali posizionati e chiusi", "PGI 3.12.7")
def c_rooms(c):
    for room in instances(doc, [BuiltInCategory.OST_Rooms]):
        if room.Location is None:
            c.add(room_label(room) + " non posizionato", room.Id)
        elif room.Area <= 0:
            c.add(room_label(room) + " non chiuso o ridondante", room.Id)


@check("07.2", "Numeri locale univoci", "PGI 3.11.12")
def c_room_numbers(c):
    numbers = {}
    for room in instances(doc, [BuiltInCategory.OST_Rooms]):
        numbers.setdefault(
            room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString(), []
        ).append(room.Id)
    for number, ids in sorted(numbers.items()):
        if len(ids) > 1:
            c.add("Numero `%s` usato da %d locali" % (number, len(ids)), ids)


@check("07.3", "Locali: UTI_TE_Piano coerente con il livello", "PGI 3.11.12")
def c_room_floor(c):
    for room in instances(doc, [BuiltInCategory.OST_Rooms]):
        state, value = param_state(room, "UTI_TE_Piano")
        if state != "value":
            continue
        m = R.RE_LEVEL.match(level_name(doc, room) or "")
        if m and value not in (m.group(2), R.wbs_l6_from_level(m.group(0))):
            c.add(
                "%s: UTI_TE_Piano `%s`, livello `%s`"
                % (room_label(room), value, m.group(2)),
                room.Id,
                WARN,
            )


# ---------------------------------------------------------------- 08 modelling rules


@check(
    "08.1",
    "Muri con vincolo superiore non collegato",
    "PGI 3.11.1 (ammesso per parapetti)",
)
def c_unconnected_walls(c):
    ids = []
    for w in instances(doc, [BuiltInCategory.OST_Walls]):
        p = w.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)
        if p is not None and p.AsElementId() == ElementId.InvalidElementId:
            ids.append(w.Id)
    if ids:
        c.add("%d muri con Top Constraint Unconnected" % len(ids), ids, WARN)


@check("08.2", "Pavimenti con shape editing (rampe: punti omogenei)", "riunione 21/09")
def c_floor_shape(c):
    for f in instances(doc, [BuiltInCategory.OST_Floors]):
        try:
            editor = f.GetSlabShapeEditor()
            if editor.IsEnabled:
                c.add(
                    "Pavimento `%s`: %d vertici modificati, verificare che non ci siano triangolazioni"
                    % (ename(f), editor.SlabShapeVertices.Size),
                    f.Id,
                    WARN,
                )
        except Exception:
            pass


@check("08.3", "Categoria Roof non ammessa (rifare come Pavimento)", "riunione 21/09")
def c_roofs(c):
    for t, n in used_types(doc, R.FORBIDDEN_CATS):
        c.add("`%s` (%d elementi)" % (ename(t), n), t.Id)


# ---------------------------------------------------------------- 09 quality and cleanup


@check("09.1", "Elementi duplicati sovrapposti", "PGI 3.14.2.1 LV1")
def c_duplicates(c):
    for fm in doc.GetWarnings():
        if "identical instances" in fm.GetDescriptionText():
            c.add(fm.GetDescriptionText(), list(fm.GetFailingElements()))


@check("09.2", "Warning Revit", "PGI 3.14.2.1 LV1")
def c_warnings(c):
    groups = {}
    for fm in doc.GetWarnings():
        groups.setdefault(fm.GetDescriptionText(), []).extend(fm.GetFailingElements())
    for text, ids in sorted(groups.items(), key=lambda kv: -len(kv[1])):
        c.add("%s (%d elementi)" % (text.split("\n")[0], len(ids)), ids, WARN)


@check("09.3", "Famiglie in-place (vietate)", "PGI 3.11.1; riunione 21/09")
def c_in_place(c):
    for fam in FilteredElementCollector(doc).OfClass(Family):
        if fam.IsInPlace:
            cat = fam.FamilyCategory.Name if fam.FamilyCategory else "-"
            c.add(
                "`%s` (%s): sostituire con famiglia caricabile" % (fam.Name, cat),
                fam.Id,
            )


@check("09.4", "DWG importati invece che collegati", "PGI 3.11.1")
def c_dwg_import(c):
    for ii in FilteredElementCollector(doc).OfClass(ImportInstance):
        if not ii.IsLinked:
            c.add("Import CAD `%s`" % ename(doc.GetElement(ii.GetTypeId())), ii.Id)


@check("09.5", "Elementi non utilizzati eliminabili", "PGI 3.14.2.1 LV1 (pulizia)")
def c_unused(c):
    counts = {}
    for element_id in doc.GetUnusedElements(HashSet[ElementId]()):
        element = doc.GetElement(element_id)
        cat = (
            element.Category.Name
            if element is not None and element.Category
            else "(senza categoria)"
        )
        counts[cat] = counts.get(cat, 0) + 1
    for cat, n in sorted(counts.items(), key=lambda kv: -kv[1]):
        c.add("%s: %d" % (cat, n), severity=WARN)


@check("09.6", "Gruppi modello", "PGI 3.11.1")
def c_groups(c):
    n = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_IOSModelGroups)
        .WhereElementIsNotElementType()
        .GetElementCount()
    )
    if n:
        c.add("%d istanze di gruppi modello" % n, severity=WARN)


# ---------------------------------------------------------------- 10 views


def is_internal_view(v):
    """Browser/system views and the revision schedules of the title blocks are not user views."""
    if v.ViewType in (ViewType.ProjectBrowser, ViewType.SystemBrowser, ViewType.Internal):
        return True
    return isinstance(v, ViewSchedule) and v.IsTitleblockRevisionSchedule


def real_views():
    return [
        v
        for v in FilteredElementCollector(doc).OfClass(View)
        if not v.IsTemplate and not is_internal_view(v)
    ]


@check("10.1", "Nome modelli di vista TIPO_TEMA_SCALA_DESCRIZIONE", "PGI 3.11.10")
def c_view_templates(c):
    for v in FilteredElementCollector(doc).OfClass(View):
        if v.IsTemplate and not R.RE_VIEW_TEMPLATE.match(v.Name):
            c.add("Modello di vista `%s`" % v.Name, v.Id, WARN)


@check(
    "10.2",
    "Nome viste in tavola (PA100-L01a.XX, SE, PR, AP)",
    "PGI 3.11.9.2 (mandatoria per le viste in tavola)",
)
def c_sheet_views(c):
    for v in real_views():
        p = v.get_Parameter(BuiltInParameter.VIEWER_SHEET_NUMBER)
        on_sheet = p is not None and (p.AsString() or "").strip() not in ("", "---")
        if not on_sheet or v.ViewType in (
            ViewType.Schedule,
            ViewType.Legend,
            ViewType.DraftingView,
        ):
            continue
        if not (R.RE_VIEW_SHEET.match(v.Name) or R.RE_VIEW_3D.match(v.Name)):
            c.add(
                "Vista in tavola `%s` (tavola %s)" % (v.Name, p.AsString()), v.Id, WARN
            )


@check("10.3", "Tema della vista compilato (WIP/COO/TAV)", "PGI 3.11.9.1")
def c_view_theme(c):
    empty, wrong = [], {}
    for v in real_views():
        if v.ViewType in (ViewType.DrawingSheet,):
            continue
        state, value = param_state(v, R.P_TEMA)
        if state == "missing":
            c.note = "Parametro `%s` non associato alle viste" % R.P_TEMA
            return
        if state == "empty":
            empty.append(v.Id)
        elif value not in R.VIEW_TEMI:
            wrong.setdefault(value, []).append(v.Id)
    if empty:
        c.add("%d viste senza Tema" % len(empty), empty, WARN)
    for value, ids in sorted(wrong.items()):
        c.add(
            "Tema `%s` non previsto (WIP/COO/TAV) su %d viste" % (value, len(ids)),
            ids,
            WARN,
        )


# ---------------------------------------------------------------- report


def select_link(ids):
    ids = [i for i in ids if i is not None][:MAX_SELECT]
    if not ids:
        return ""
    return output.linkify(ids, title="seleziona %d" % len(ids))


def save_csv():
    """One row per finding: evidence of the LV1 check to attach to the model shares."""
    path = os.path.join(
        os.path.expanduser("~"),
        "Documents",
        "BMC_ModelChecker_%s_%s.csv"
        % (doc.Title, datetime.datetime.now().strftime("%Y%m%d_%H%M")),
    )
    with codecs.open(path, "w", encoding="utf-8-sig") as f:
        f.write("ID;Controllo;Esito;Fonte;Voce;Elementi\n")
        for c in checks:
            rows = c.items or [(None, c.note or "")]
            for element_id, text in rows:
                ids = (
                    element_id
                    if isinstance(element_id, list)
                    else ([element_id] if element_id else [])
                )
                f.write(
                    "%s;%s;%s;%s;%s;%s\n"
                    % (
                        c.cid,
                        c.title.replace(";", ","),
                        LABEL[c.status],
                        c.source.replace(";", ","),
                        # one line per finding: Revit warning texts contain line breaks
                        " ".join(text.replace(";", ",").replace("`", "").split()),
                        " ".join(str(idv(i)) for i in ids[:MAX_SELECT]),
                    )
                )
    return path


def print_report(csv_path):
    output.print_md("# BMC - Model Checker ARC (PGI R06)")
    output.print_md("Modello: **%s** - regole PGI R06 / PIR R06 / WBS R06." % doc.Title)
    totals = {OK: 0, WARN: 0, ERR: 0, FAIL: 0}
    rows = []
    for c in checks:
        totals[c.status] += 1
        rows.append([c.cid, c.title, LABEL[c.status], len(c.items), c.source])
    output.print_md(
        "**%d controlli** - OK %d - AVVISO %d - ERRORE %d - ERRORE SCRIPT %d"
        % (len(checks), totals[OK], totals[WARN], totals[ERR], totals[FAIL])
    )
    output.print_md("Report CSV: `%s`" % csv_path)
    output.print_table(
        table_data=rows, columns=["ID", "Controllo", "Esito", "Voci", "Fonte"]
    )
    for c in checks:
        if c.status == OK and not c.note:
            continue
        output.print_md("---")
        output.print_md("### %s - %s: %s" % (c.cid, c.title, LABEL[c.status]))
        output.print_md("*Fonte: %s*" % c.source)
        if c.note:
            output.print_md(c.note)
        all_ids = []
        for element_id, _ in c.items:
            if isinstance(element_id, list):
                all_ids.extend(element_id)
            elif element_id is not None:
                all_ids.append(element_id)
        if len(all_ids) > 1:
            output.print_md("Tutti gli elementi: " + select_link(all_ids))
        for element_id, text in c.items[:MAX_ROWS]:
            if isinstance(element_id, list):
                output.print_md("- %s %s" % (text, select_link(element_id)))
            elif element_id is not None:
                output.print_md("- %s %s" % (text, output.linkify(element_id)))
            else:
                output.print_md("- %s" % text)
        if len(c.items) > MAX_ROWS:
            output.print_md(
                "- ... altre %d voci (tutte nel CSV)" % (len(c.items) - MAX_ROWS)
            )


output.resize(1300, 900)
for runner in runners:
    runner()
print_report(save_csv())
