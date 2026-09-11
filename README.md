# oa — Office Automation CLI

A Windows CLI tool that automates Microsoft Office (PowerPoint + Excel) via the COM API. Built in Rust for speed and safety.

Takes a PowerPoint template with linked OLE objects and Excel data, then updates tables, charts, delta indicators, and color coding — producing a fully populated report in seconds.

## Features

- **Update pipeline** — Re-link OLE objects, populate tables, swap delta arrows, apply color coding, update charts
- **Batch processing** — TOML runfiles for processing dozens of files in one command
- **Validation** — Cell-by-cell check of PPT values against Excel source data
- **ZIP pre-processing** — Rewrite OLE/chart paths and chart data directly in PPTX XML (100x faster than COM for links)
- **Inspection** — Read-only analysis of PPTX shape inventory with per-slide breakdown

## Installation

Requires Windows with Microsoft Office (PowerPoint + Excel) installed.

### One-liner (no Rust needed)

**PowerShell:**
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://raw.githubusercontent.com/albertxli/office-automation/main/install.ps1 | iex"
```

Downloads `oa.exe`, installs to `%LOCALAPPDATA%\oa`, and adds it to your PATH automatically. Restart your terminal after install.

Or download manually from [GitHub Releases](https://github.com/albertxli/office-automation/releases/latest).

### Via crates.io (requires Rust)

```bash
cargo install office-automation
```

### Build from source

```bash
git clone https://github.com/albertxli/office-automation.git
cd office-automation
cargo install --path .
```

## Quick Start

```bash
# Update a presentation with new Excel data
oa update template.pptx -e data.xlsx -o report.pptx

# Batch process from a runfile
oa run batch.toml

# Validate output against Excel
oa check report.pptx -e data.xlsx

# Inspect a PPTX file
oa info template.pptx

# Per-slide shape breakdown
oa info -v template.pptx
```

## Commands

| Command | Description |
|---------|-------------|
| `oa update` | Run the update pipeline on PPTX files |
| `oa run` | Execute a TOML runfile for batch processing |
| `oa check` | Validate PPT values against Excel source |
| `oa info` | Inspect a PPTX file (read-only) |
| `oa diff` | Compare two PPTX files side by side |
| `oa config` | Show all config keys and defaults |
| `oa clean` | Kill zombie Office processes |

See [API.md](API.md) for full command reference with all options and examples.

## Runfile Example

```toml
data_path = "../data"
default_output = "output/{name}.pptx"

[templates]
t1 = "templates/template.pptx"

[[job]]
name = "Australia"
template = "t1"
data = "tracking_australia.xlsx"

[[job]]
name = "Japan"
template = "t1"
data = "tracking_japan.xlsx"
```

```bash
oa run batch.toml
```

See [example_runfile.toml](example_runfile.toml) for a complete example.

## Pipeline Steps

The update pipeline runs these steps in order:

| Step | Description |
|------|-------------|
| **Links** | Re-point OLE links to the new Excel file |
| **Tables** | Populate PPT table cells from Excel ranges |
| **Deltas** | Swap delta indicator arrows based on value sign, with optional dead-band thresholds (see below) |
| **Coloring** | Apply sign-based color coding to _ccst tables |
| **Charts** | Rebuild each chart's data cache from Excel and re-point links; blank cells draw nothing, a real 0 draws a zero bar; series formulas that name a workbook (`[book.xlsx]Sheet!Range`) are normalised and reported as a warning |

Steps can be selectively run or skipped:

```bash
oa update report.pptx -e data.xlsx --steps tables,charts
oa update report.pptx -e data.xlsx --skip deltas
```

## Shape Naming Conventions

The pipeline identifies shapes by name prefix:

| Prefix | Type | Description |
|--------|------|-------------|
| `ntbl_` | Normal table | Updates cell text, preserves formatting |
| `htmp_` | Heatmap table | Applies 3-color scale from Excel |
| `trns_` | Transposed table | Swaps rows/columns from Excel |
| `delt_` | Delta indicator | Arrow swapped by value sign |
| `delt<N>_` | Delta indicator, set N | Same as `delt_`, but copies from the `tmpl<N>_delta_*` templates |
| `_ccst` | Color-coded | Cells colored positive/negative/neutral |

### Multiple delta template sets

Delta arrows are replaced by copying a template shape from slide 1. When a deck needs
more than one arrow style (for example arrows with text on most slides and text-free
arrows on a few), number the set on both sides:

| Set | Shape prefix | Templates on slide 1 |
|-----|--------------|----------------------|
| 1 | `delt_` (or `delt1_`) | `tmpl_delta_pos`, `tmpl_delta_neg`, `tmpl_delta_none` |
| 2 | `delt2_` | `tmpl2_delta_pos`, `tmpl2_delta_neg`, `tmpl2_delta_none` |
| N | `delt<N>_` | `tmpl<N>_delta_pos`, `tmpl<N>_delta_neg`, `tmpl<N>_delta_none` |

The OLE name and the `_pos/_neg/_none` suffix work exactly as for `delt_`, e.g.
`delt2_Rev_DE_pos`. If a set's templates are missing, that set is skipped with a
warning and other sets still update. `oa info` lists every set found and its templates;
`oa check` validates all sets.

### Delta thresholds

By default a delta is `pos` when the cell is `> 0`, `neg` when `< 0`, else `none`. Net-opinion
cells are built from rounded percentages, so tiny deltas are noise. A dead band turns them into
`none`: with threshold `t`, `value >= t` is `pos`, `value <= -t` is `neg`, anything strictly in
between is `none`.

Thresholds are config keys. `delta.threshold` applies to every delta; `delta.threshold.<token>`
applies to deltas whose paired OLE object name contains `<token>` as a whole word (`_` counts as
a word boundary, so `globalnet` covers `globalnet_pet` and `globalnet_dig` but not
`globalnetwork`). When several tokens match, the longest wins.

```bash
# CLI: one --set per key
oa update report.pptx -e france.xlsx --set delta.threshold.globalnet=0.02 --set delta.threshold.marketnet=0.05
```

```toml
# TOML runfile — keys contain dots, so quote them
[config]
"delta.threshold.globalnet" = 0.02
"delta.threshold.marketnet" = 0.05
```

```python
# Python runfile
config = {"delta.threshold.globalnet": 0.02, "delta.threshold.marketnet": 0.05}
```

Values are always decimals, read numerically from Excel: a cell typed as a percentage and
showing `2%` is `0.02`. With the two thresholds above:

| Cell shows | Read as | `globalnet` (0.02) | `marketnet` (0.05) |
|-----------|---------|--------------------|--------------------|
| -3% | -0.03 | neg | none |
| -2% | -0.02 | neg (boundary is inclusive) | none |
| -5% | -0.05 | neg | neg |
| -1% | -0.01 | none | none |

A text or blank cell under a threshold prints a warning and sets the delta to `none`. Deltas with
threshold `0` (the default) keep the original sign test unchanged. Run `oa check` with the same
`--set` values, otherwise it reports the dead-band deltas as mismatches. `-v` shows the threshold
and token used per delta (`-0.02 → neg · thr 0.02 via globalnet`).

## Performance

| Scenario | Time |
|----------|------|
| Single 68-slide PPTX (155 OLE, 257 charts) | ~6s |
| Batch 26 files via `oa run` | ~2m 36s |
| ZIP pre-relink (411 links) | 0.1s |

## License

[MIT](LICENSE)
