"""Independent Aspen Plus V15 COM smoke case, with save/reopen verification.

Usage: python smoke_background.py --out F:/AspenRuns/smoke_001
Use a NEW output directory. No existing document is attached or overwritten.
IDEAL is only a connectivity/automation test, not a process design choice.
"""
import argparse
import hashlib
import json
import math
from pathlib import Path
import re
import sys
import time
import traceback

import pythoncom
import win32com.client


INPUT = """IN-UNITS MET SHORT-LENGTH=mm

DATABANKS 'APV150 PURE41' / 'APV150 AQUEOUS' / 'APV150 SOLIDS'

PROP-SOURCES 'APV150 PURE41' / 'APV150 AQUEOUS' / 'APV150 SOLIDS'

COMPONENTS
    PHENOL PHENOL /
    H2O WATER /
    H2 HYDROGEN

PROPERTIES IDEAL

FLOWSHEET
    BLOCK CHECK IN=FEED OUT=EFF

STREAM FEED
    SUBSTREAM MIXED TEMP=473 <K> PRES=40 <bar> MASS-FLOW=102 <kg/hr>
    MASS-FLOW PHENOL 1 / H2O 100 / H2 1

BLOCK CHECK HEATER
    PARAM TEMP=473 <K> PRES=40 <bar> NPHASE=2
"""


def node(app, path):
    found = app.Tree.FindNode(path)
    if found is None:
        raise RuntimeError(f"Missing Aspen node: {path}")
    return found


def quantity(app, path):
    item = node(app, path)
    value = item.Value
    if value is None or not math.isfinite(float(value)):
        raise RuntimeError(f"Missing/nonfinite result: {path} = {value!r}")
    return {"value": float(value), "unit": str(item.UnitString)}


def state(app):
    # V15: HAP_COMPSTATUS attribute 12; confirm via typelib on other versions.
    flags = int(node(app, r"\Data").AttributeValue(12))
    return {
        "HAP_COMPSTATUS": flags,
        "input_incomplete": bool(flags & 64),
        "engine_ready": bool(app.Engine.Ready),
        "next_incomplete": node(app, r"\Data").NextIncomplete(0),
    }


def start(path, progid):
    app = win32com.client.DispatchEx(progid)
    try:
        app.SuppressDialogs = 1
        if path.suffix.lower() == ".bkp":
            app.InitFromArchive2(str(path))
        else:
            app.InitFromFile2(str(path))
        app.Visible = False
        return app
    except Exception:
        try:
            app.Quit()
        except Exception:
            pass
        raise


def close(app):
    if app is not None:
        try:
            app.Quit()
        except Exception as exc:
            print(f"Cleanup: {exc}", file=sys.stderr, flush=True)


def calculate(app, out, label):
    before = state(app)
    if before["input_incomplete"] or not before["engine_ready"]:
        raise RuntimeError(f"Input check failed: {before}")
    print(f"{label}: imported; running", flush=True)
    # Synchronous COM call: enforce a wall-clock deadline in an outer supervisor.
    app.Engine.Run2(False)
    app.Export(13, str(out / f"{label}.his"))
    app.Export(4, str(out / f"{label}_solved.inp"))
    history = (out / f"{label}.his").read_text(encoding="cp1252", errors="replace")
    errors = re.findall(
        r"^\s*\*{2,3}\s+(?:SEVERE ERROR|TERMINAL ERROR|ERROR)[^\n]*",
        history, re.M,
    )
    after = state(app)
    if errors or after["input_incomplete"]:
        raise RuntimeError(f"Solver/input failure: {errors}; {after}")
    feed = quantity(app, r"\Data\Streams\FEED\Output\MASSFLMX\MIXED")
    eff = quantity(app, r"\Data\Streams\EFF\Output\MASSFLMX\MIXED")
    if feed["unit"] != "kg/hr" or eff["unit"] != "kg/hr":
        raise RuntimeError(f"Unexpected mass units: {feed}, {eff}")
    residual = eff["value"] - feed["value"]
    if abs(feed["value"] - 102) > 1e-6 or abs(residual) > 1e-6:
        raise RuntimeError(f"Mass check failed: feed={feed}, eff={eff}")
    return {
        "before": before, "after": after, "feed": feed, "eff": eff,
        "mass_residual_kg_h": residual, "solver_errors": errors,
        "warning_lines": [s for s in history.splitlines() if "WARNING" in s.upper()],
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", required=True)
    parser.add_argument("--progid", default="Apwn.Document.41.0")
    args = parser.parse_args()
    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=False)
    inp = out / "smoke.inp"
    inp.write_text(INPUT, encoding="ascii")
    report = {"status": "started", "progid": args.progid, "input": str(inp),
              "input_sha256": hashlib.sha256(inp.read_bytes()).hexdigest()}
    app = None
    started = time.monotonic()
    pythoncom.CoInitialize()
    try:
        app = start(inp, args.progid)
        report["first_run"] = calculate(app, out, "first")
        archive = out / "smoke.bkp"
        document = out / "smoke.apwz"
        app.WriteArchive2(str(archive), 1)
        app.SaveAs2(str(document), True)
        for path in (archive, document):
            if not path.is_file() or not path.stat().st_size:
                raise RuntimeError(f"Save failed: {path}")
        close(app)
        app = None
        for label, path in (("reopened_document", document), ("reopened_archive", archive)):
            app = start(path, args.progid)
            report[label] = calculate(app, out, label)
            if abs(report[label]["eff"]["value"] - report["first_run"]["eff"]["value"]) > 1e-6:
                raise RuntimeError(f"Reopened result mismatch: {label}")
            close(app)
            app = None
        report["status"] = "passed"
    except Exception:
        report["status"] = "failed"
        report["traceback"] = traceback.format_exc()
    finally:
        close(app)
        pythoncom.CoUninitialize()
        report["elapsed_seconds"] = time.monotonic() - started
        (out / "report.json").write_text(
            json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2, default=str), flush=True)
    return 0 if report["status"] == "passed" else 1


if __name__ == "__main__":
    sys.exit(main())
