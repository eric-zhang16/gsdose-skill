#!/usr/bin/env python3
"""
gsDose Report Template - generates gsdose_report.docx from gsdose_results.json.
Pure template: ZERO hardcoded design values. Every number comes from the JSON.

Usage:
    python gsdose_report_template.py <output_dir>

The script expects:
    <output_dir>/gsdose_results.json
    (written by compute_boundary.R; compute_test.R merges its results into the
    same file under each endpoint's "test" field)

It produces:
    <output_dir>/gsdose_report.docx

Sections:
    1. Design Assumptions
    2. Interim/Final Analysis Plan
    3. Efficacy Boundaries
    4. Combination Test and Closed Testing   (only when an endpoint has test results)
"""

import json
import os
import sys
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


# -- Parse CLI ----------------------------------------------------------------

out_dir = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
json_path = os.path.join(out_dir, "gsdose_results.json")
report_path = os.path.join(out_dir, "gsdose_report.docx")

with open(json_path) as f:
    r = json.load(f)

meta = r.get("meta", {})
endpoints = r.get("endpoints", [])


# -- Helpers ------------------------------------------------------------------

def pct(x, dec=1):
    v = x * 100 if x <= 1 else x
    return f"{v:.{dec}f}%"

def fmt(x, dec=3):
    try:
        return f"{float(x):.{dec}f}"
    except (TypeError, ValueError):
        return str(x)

def add_heading(doc, text, level=1):
    return doc.add_heading(text, level=level)

def add_para(doc, text="", bold=False, italic=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    return p

def add_table(doc, headers, rows):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hdr = t.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = h
        for run in cell.paragraphs[0].runs:
            run.bold = True
    for ri, row in enumerate(rows):
        tr = t.rows[ri + 1]
        for ci, val in enumerate(row):
            tr.cells[ci].text = "" if val is None else str(val)
    return t

def m(key, default=""):
    return meta.get(key, default)

def spending_label(sf, sfpar):
    if sf == "OF":
        return "O'Brien-Fleming (Lan-DeMets)"
    if sf == "HSD":
        return f"Hwang-Shih-DeCani (gamma = {sfpar})"
    return str(sf)


# -- Build document -----------------------------------------------------------

doc = Document()

title_p = doc.add_paragraph()
title_run = title_p.add_run(
    f"Seamless Phase 2/3 Dose-Selection Design Report: {m('disease', '')}"
)
title_run.bold = True
title_run.font.size = Pt(16)
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.add_paragraph()


# ==============================================================================
# Section 1: Design Assumptions
# ==============================================================================

add_heading(doc, "1. Design Assumptions", 1)

rows = [
    ["Disease / Setting", str(m("disease"))],
    ["Endpoints", str(m("endpoints"))],
    ["Dose arms at start (plus control)", str(m("k_dose_arms"))],
    ["Doses selected at dose selection", "1 (single dose carried forward)"],
    ["Randomization (dose : control)", str(m("randomization_ratio"))],
    ["Patients per arm before dose selection (n1)", str(m("n1_per_arm"))],
    ["Patients per arm after dose selection (n2)", str(m("n2_per_arm"))],
]

if m("total_N"):
    rows.append(["Total N", str(m("total_N"))])
if m("ctrl_median"):
    rows.append(["Control median survival", f"{m('ctrl_median')} months"])
if m("hr"):
    rows.append(["Assumed hazard ratio (all dose arms)", str(m("hr"))])
if m("dropout_rate") not in ("", None):
    rows.append(["Annual dropout", str(m("dropout_rate"))])
if m("enroll_duration"):
    rows.append(["Enrollment duration", f"{m('enroll_duration')} months"])

for ep in endpoints:
    label = f"Alpha (one-sided), {ep.get('name', '')}"
    rows.append([label, str(ep.get("alpha", ""))])
    rows.append([
        f"Efficacy spending, {ep.get('name', '')}",
        spending_label(ep.get("sf"), ep.get("sfpar")),
    ])
    fut = ep.get("futility") or {}
    if fut.get("enabled"):
        rows.append([
            f"Futility (lower) bound, {ep.get('name', '')}",
            (f"Null-based spending of {fut.get('total')} total, "
             f"{spending_label(fut.get('sf'), fut.get('sfpar'))}"),
        ])
    else:
        rows.append([f"Futility bound, {ep.get('name', '')}", "None (efficacy only)"])
    hrow = ep.get("h")
    if hrow:
        h1 = hrow[0][0] if isinstance(hrow[0], list) else hrow[0]
        h2 = hrow[0][1] if isinstance(hrow[0], list) else hrow[1]
        rows.append([
            f"Cohort weights h (cohort 1, cohort 2), {ep.get('name', '')}",
            f"({fmt(h1)}, {fmt(h2)})",
        ])

add_table(doc, ["Parameter", "Value"], rows)
add_para(
    doc,
    "Cohort 1 comprises patients enrolled before dose selection (all dose arms "
    "and control); cohort 2 comprises patients enrolled after dose selection "
    "(selected dose and control). Event counts below refer to the selected dose "
    "arm and control only. The hazard ratio assumption applies to all dose arms "
    "and affects efficiency only, not type I error control, because the "
    "combination weights are pre-specified.",
    italic=True,
)
doc.add_paragraph()


# ==============================================================================
# Section 2: Interim/Final Analysis Plan
# ==============================================================================

add_heading(doc, "2. Interim/Final Analysis Plan", 1)

for ep in endpoints:
    name = ep.get("name", "")
    stages = ep.get("stages", [])
    s_total = ep.get("s", len(stages))
    if len(endpoints) > 1:
        add_heading(doc, name, 2)

    add_para(
        doc,
        (f"{name} is tested at {s_total} analyses (dose-selection interim "
         f"excluded), with {ep.get('planD')} events planned from the selected "
         f"dose arm and control at the final analysis."),
    )

    has_time = any("expected_time" in st for st in stages)
    headers = ["Analysis", "Events (both cohorts)", "Cohort 1 events",
               "Cohort 2 events", "Information fraction"]
    if has_time:
        headers.append("Expected time (mo)")
    trows = []
    for st in stages:
        label = f"IA{st['stage']}" if st["stage"] < s_total else "FA"
        row = [label, st.get("d"), st.get("d1"), st.get("d2"),
               pct(st.get("info_frac", 0))]
        if has_time:
            row.append(st.get("expected_time", ""))
        trows.append(row)
    add_table(doc, headers, trows)
    doc.add_paragraph()


# ==============================================================================
# Section 3: Efficacy Boundaries
# ==============================================================================

add_heading(doc, "3. Efficacy Boundaries", 1)

for ep in endpoints:
    name = ep.get("name", "")
    stages = ep.get("stages", [])
    s_total = ep.get("s", len(stages))
    if len(endpoints) > 1:
        add_heading(doc, name, 2)

    add_para(
        doc,
        (f"Efficacy boundaries for the inverse normal combination test, from "
         f"spending alpha = {ep.get('alpha')} with "
         f"{spending_label(ep.get('sf'), ep.get('sfpar'))} at the observed/planned "
         f"information fractions."),
    )

    fut = ep.get("futility") or {}
    has_lower = fut.get("enabled") and any("z_lower" in st for st in stages)
    headers = ["Analysis", "Information fraction", "Z boundary (efficacy)",
               "Nominal one-sided p"]
    if has_lower:
        headers.append("Z boundary (futility)")
    trows = []
    for st in stages:
        label = f"IA{st['stage']}" if st["stage"] < s_total else "FA"
        row = [label, pct(st.get("info_frac", 0)), fmt(st.get("z_upper"), 3),
               fmt(st.get("nominal_p_upper"), 4)]
        if has_lower:
            row.append(fmt(st.get("z_lower"), 3))
        trows.append(row)
    add_table(doc, headers, trows)

    if has_lower:
        add_para(
            doc,
            ("The futility (lower) boundary spends a total probability of "
             f"{fut.get('total')} under the null hypothesis across stages using "
             f"{spending_label(fut.get('sf'), fut.get('sfpar'))}; it is a "
             "null-based spending rule, not a beta-spending rule."),
            italic=True,
        )
    doc.add_paragraph()


# ==============================================================================
# Section 4: Combination Test and Closed Testing (workflow B, optional)
# ==============================================================================

tested = [ep for ep in endpoints if ep.get("test")]
if tested:
    add_heading(doc, "4. Combination Test and Closed Testing", 1)

    for ep in tested:
        t = ep["test"]
        if len(tested) > 1 or len(endpoints) > 1:
            add_heading(doc, ep.get("name", ""), 2)

        add_para(
            doc,
            (f"Analysis at testing stage {t.get('stage')} for {ep.get('name', '')}: "
             f"dose {t.get('sid')} was selected at dose selection. Each intersection "
             f"hypothesis containing the selected dose is tested with the inverse "
             f"normal combination of the cohort-1 Dunnett statistic and the cohort-2 "
             f"log-rank statistic, against the efficacy boundary "
             f"Z = {fmt(t.get('zbound'), 3)} computed at the actual event counts."),
        )

        headers = ["Intersection", "Doses included", "Dunnett Z (cohort 1)",
                   "Cohort-2 Z", "Combination Z", f"Reject (Z > {fmt(t.get('zbound'), 3)})"]
        trows = []
        for it in t.get("intersections", []):
            doses = it.get("doses")
            if isinstance(doses, list):
                doses_str = ", ".join(str(x) for x in doses)
            else:
                doses_str = str(doses)
            trows.append([
                it.get("hypothesis", ""),
                doses_str,
                fmt(it.get("dunnett_z1"), 3),
                fmt(it.get("logrank_z2"), 3),
                fmt(it.get("combo_z"), 3),
                "Yes" if it.get("reject") else "No",
            ])
        add_table(doc, headers, trows)

        add_para(doc, t.get("decision", ""), bold=True)

        fut = t.get("futility")
        if fut:
            crossed = "crossed" if fut.get("crossed") else "did not cross"
            add_para(
                doc,
                (f"Futility check (informational): the selected dose's combination "
                 f"Z of {fmt(fut.get('individual_combo_z'), 3)} {crossed} the futility "
                 f"boundary Z = {fmt(fut.get('z_lower'), 3)}."),
                italic=True,
            )
        doc.add_paragraph()


# -- Save ----------------------------------------------------------------------

doc.save(report_path)
print(f"WROTE: {report_path}")
