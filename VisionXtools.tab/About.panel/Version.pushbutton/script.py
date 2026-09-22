# -*- coding: utf-8 -*-
"""Show the installed VisionXtools version, its date and what changed."""

__title__ = "Version"
__author__ = "Luca Rosati"
__context__ = "zero-doc"

import io
import os
import re
import time

from pyrevit import script

output = script.get_output()

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
CHANGELOG = os.path.join(ROOT, "CHANGELOG.md")
SECTION = re.compile(r"^##\s+(\S+)\s+[-–—]\s+(.+?)\s*$", re.MULTILINE)


def installed_commit():
    """Short sha and date of the installed commit, read straight from .git.

    ponytail: parses .git/logs/HEAD so the button works without the git CLI;
    returns None on a copy installed without git.
    """
    log = os.path.join(ROOT, ".git", "logs", "HEAD")
    if not os.path.isfile(log):
        return None
    try:
        with io.open(log, encoding="utf-8", errors="replace") as fh:
            lines = [l for l in fh.read().splitlines() if l.strip()]
        head = lines[-1].split("\t")[0].split()
        return head[1][:7], time.strftime(
            "%d/%m/%Y %H:%M", time.localtime(int(head[-2]))
        )
    except Exception:
        return None


def sections(text):
    """[(version, date, body)] in file order, newest first."""
    found = list(SECTION.finditer(text))
    out = []
    for i, m in enumerate(found):
        end = found[i + 1].start() if i + 1 < len(found) else len(text)
        out.append((m.group(1), m.group(2), text[m.end() : end].strip()))
    return out


if not os.path.isfile(CHANGELOG):
    output.print_md(
        "# VisionXtools\n\n**CHANGELOG.md non trovato** in `{}`.\n\n"
        "La copia installata e' incompleta: reinstallala con "
        "`Installa_VisionXtools.bat`.".format(ROOT)
    )
else:
    with io.open(CHANGELOG, encoding="utf-8", errors="replace") as fh:
        entries = sections(fh.read())

    if not entries:
        output.print_md("# VisionXtools\n\nNessuna versione elencata nel CHANGELOG.md.")
    else:
        version, date, body = entries[0]
        output.print_md("# VisionXtools {}".format(version))
        output.print_md("**Data di rilascio: {}**".format(date))
        output.print_md("### Novita' di questa versione\n\n{}".format(body))

        commit = installed_commit()
        output.print_md("---")
        if commit:
            output.print_md("Copia installata: commit `{}` del {}".format(*commit))
        else:
            output.print_md(
                "Copia installata senza git: non aggiornabile con "
                "About > Update, va reinstallata con "
                "`Installa_VisionXtools.bat`."
            )
        output.print_md("Cartella: `{}`".format(ROOT))

        if len(entries) > 1:
            output.print_md("## Versioni precedenti")
            for version, date, body in entries[1:]:
                output.print_md("### {} - {}\n\n{}".format(version, date, body))
