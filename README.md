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

```bash
cargo install office-automation
```

Or build from source:

```bash
git clone https://github.com/user/office-automation.git
cd office-automation
cargo build --release
```

The binary is `oa.exe` in `target/release/`.

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
| **Deltas** | Swap delta indicator arrows based on value sign |
| **Coloring** | Apply sign-based color coding to _ccst tables |
| **Charts** | Update chart data links |

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
| `_ccst` | Color-coded | Cells colored positive/negative/neutral |

## Performance

| Scenario | Time |
|----------|------|
| Single 68-slide PPTX (155 OLE, 257 charts) | ~6s |
| Batch 26 files via `oa run` | ~2m 36s |
| ZIP pre-relink (411 links) | 0.1s |

## License

[MIT](LICENSE)
