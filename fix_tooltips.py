#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_tooltips.py - Popola il campo tooltip in bundle.yaml per ogni pushbutton
leggendo il docstring da script.py. Esegui con Python 3 standard (non IronPython).

Usage:
    python fix_tooltips.py --dry-run   # preview senza scrivere
    python fix_tooltips.py --verbose   # output dettagliato
    python fix_tooltips.py             # applica le modifiche
"""

import ast
import os
import re
import sys
import argparse
import textwrap

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

EXTENSION_ROOT = os.path.dirname(os.path.abspath(__file__))

_DOCSTRING_RE = re.compile(r'^\s*(?:#[^\n]*\n)*\s*"""(.*?)"""', re.DOTALL)


def extract_docstring(script_path):
    with open(script_path, "r", encoding="utf-8", errors="replace") as fh:
        source = fh.read()
    try:
        tree = ast.parse(source, filename=script_path)
        doc = ast.get_docstring(tree, clean=True)
        if doc:
            return doc
    except SyntaxError:
        pass
    # Fallback: regex on raw source for scripts with invalid escape sequences
    m = _DOCSTRING_RE.match(source)
    if m:
        return m.group(1).strip()
    return None


def yaml_scalar(text):
    YAML_SPECIAL = set(":{}\n[]#&*?|-<>=!%@`")
    if "\n" in text:
        indented = textwrap.indent(text, "  ").rstrip("\n") + "\n"
        return "|\n" + indented
    elif any(c in YAML_SPECIAL for c in text):
        escaped = text.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{escaped}"'
    else:
        return f'"{text}"'


def has_tooltip_key(yaml_text):
    for line in yaml_text.splitlines():
        if line.strip().startswith("tooltip:"):
            return True
    return False


def find_pushbutton_folders(root):
    for dirpath, dirnames, _ in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        for d in dirnames:
            if d.endswith(".pushbutton"):
                yield os.path.join(dirpath, d)


def process_pushbutton(folder, dry_run=False, verbose=False):
    script_path = os.path.join(folder, "script.py")
    if not os.path.isfile(script_path):
        if verbose:
            print(f"  [SKIP] No script.py")
        return "skipped"

    docstring = extract_docstring(script_path)
    if not docstring:
        print(f"  [WARN] No docstring: {script_path}")
        return "no_docstring"

    yaml_path = os.path.join(folder, "bundle.yaml")
    yml_path = os.path.join(folder, "bundle.yml")
    active_path = (
        yaml_path
        if os.path.isfile(yaml_path)
        else (yml_path if os.path.isfile(yml_path) else None)
    )

    if active_path is not None:
        with open(active_path, "r", encoding="utf-8", errors="replace") as fh:
            yaml_text = fh.read()
        if has_tooltip_key(yaml_text):
            if verbose:
                print(f"  [SKIP] tooltip già presente: {os.path.basename(active_path)}")
            return "already_has_tooltip"
        scalar = yaml_scalar(docstring)
        new_content = yaml_text.rstrip("\n") + f"\ntooltip: {scalar}\n"
        action = f"UPDATE {os.path.basename(active_path)}"
        target_path = active_path
    else:
        scalar = yaml_scalar(docstring)
        new_content = f"tooltip: {scalar}\n"
        action = "CREATE bundle.yaml"
        target_path = yaml_path

    if verbose:
        print(f"  [{action}] — {docstring[:70]}{'...' if len(docstring) > 70 else ''}")
    else:
        print(f"  [{action}]")

    if not dry_run:
        with open(target_path, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(new_content)

    return "created" if active_path is None else "updated"


def main():
    parser = argparse.ArgumentParser(
        description="Popola tooltip in bundle.yaml da docstring script.py"
    )
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--verbose", "-v", action="store_true")
    args = parser.parse_args()

    if args.dry_run:
        print("DRY RUN — nessun file verrà scritto.\n")

    counters = {
        "created": 0,
        "updated": 0,
        "already_has_tooltip": 0,
        "no_docstring": 0,
        "skipped": 0,
    }
    folders = sorted(find_pushbutton_folders(EXTENSION_ROOT))
    print(f"Trovate {len(folders)} cartelle pushbutton\n")

    for folder in folders:
        rel = os.path.relpath(folder, EXTENSION_ROOT)
        print(f"{rel}")
        status = process_pushbutton(folder, dry_run=args.dry_run, verbose=args.verbose)
        counters[status] += 1

    print("\n" + "-" * 60)
    print(f"  Creati nuovi bundle.yaml : {counters['created']}")
    print(f"  Aggiornati               : {counters['updated']}")
    print(f"  Già con tooltip          : {counters['already_has_tooltip']}")
    print(f"  Senza docstring (fix man): {counters['no_docstring']}")
    print(f"  Saltati (no script.py)   : {counters['skipped']}")


if __name__ == "__main__":
    main()
