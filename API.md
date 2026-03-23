# oa — Office Automation CLI Reference

Windows-only CLI tool that automates Microsoft Office (PowerPoint + Excel) via COM.

## Quick Start

```bash
# Update a single presentation with new Excel data
oa update report.pptx -e data.xlsx

# Update and save to a new file
oa update template.pptx -e data.xlsx -o output/report.pptx

# Inspect a PPTX file (read-only)
oa info report.pptx

# Show all config keys and defaults
oa config

# Kill zombie Office processes
oa clean -f
```

---

## Commands

### `oa update` — Run the update pipeline

The main command. Processes PPTX files by re-linking OLE objects, populating tables, swapping delta indicators, applying color coding, and updating charts.

```
oa update <FILES...> [OPTIONS]
```

**Arguments:**

| Argument | Description |
|----------|-------------|
| `<FILES>` | One or more PPTX files (glob patterns like `*.pptx` supported) |

**Options:**

| Flag | Description |
|------|-------------|
| `-e, --excel <PATH>` | Excel data file. Auto-detected from OLE links if omitted |
| `-p, --pick` | Open native file dialog to select Excel file |
| `--pair <PPT=XLSX>` | Explicit PPTX=XLSX pair (repeatable) |
| `-o, --output <PATH>` | Output file or directory. Default: in-place |
| `--steps <STEP,...>` | Run only these steps (comma-separated) |
| `--skip <STEP,...>` | Skip these steps (mutually exclusive with --steps) |
| `--set <KEY=VALUE>` | Override a config value (repeatable) |
| `--check` | Run validation against Excel after processing |
| `--dry-run` | Show what would happen without saving |
| `-v, --verbose` | Enable debug logging |
| `-q, --quiet` | Suppress all output except errors |

**Pipeline Steps** (executed in this order):

| Step | Requires Excel | Description |
|------|---------------|-------------|
| `links` | Yes | Re-point OLE links to new Excel file |
| `tables` | Yes | Populate PPT tables from Excel ranges |
| `deltas` | Yes | Swap delta indicator arrows based on sign |
| `coloring` | No | Apply sign-based color coding (_ccst shapes) |
| `charts` | Yes | Update chart data links |

**Examples:**

```bash
# Basic: update template with new data
oa update template.pptx -e quarterly_data.xlsx

# Save to output directory (original unchanged)
oa update template.pptx -e data.xlsx -o output/report.pptx

# Only update tables and charts (skip links, deltas, coloring)
oa update report.pptx -e data.xlsx --steps tables,charts

# Skip chart updates (everything else runs)
oa update report.pptx -e data.xlsx --skip charts

# Update multiple files with the same Excel
oa update "reports/*.pptx" -e data.xlsx

# Explicit pairs (different Excel per PPTX)
oa update --pair us_report.pptx=us_data.xlsx --pair mx_report.pptx=mx_data.xlsx

# Override config values
oa update report.pptx -e data.xlsx --set ccst.positive_color=#00FF00 --set links.set_manual=false

# Dry run: see what would happen without saving
oa update report.pptx -e data.xlsx --dry-run

# Update and validate results
oa update report.pptx -e data.xlsx --check

# Auto-detect Excel from OLE links in the PPTX
oa update report.pptx

# Quiet mode for scripts
oa update report.pptx -e data.xlsx -q

# Open file dialog to select Excel
oa update report.pptx --pick
```

---

### `oa info` — Inspect a PPTX file

Read-only inspection. Shows slide count, OLE links, charts, special shapes, and delta templates.

```
oa info <FILE>
```

**Example:**

```bash
oa info template.pptx
```

**Output:**

```
Presentation
  File:   template.pptx
  Slides: 30

OLE Links
  C:\Data\report.xlsx                                              86
  Total                                                            86

Charts
  Linked:   100
  Unlinked: 0

Special Shapes
  ntbl_ (normal tables):    49
  htmp_ (heatmap tables):   0
  trns_ (transposed):       31
  delt_ (delta indicators): 6
  _ccst (color-coded):      24

Delta Templates (Slide 1)
  tmpl_delta_pos                 ✓
  tmpl_delta_neg                 ✓
  tmpl_delta_none                ✓
```

---

### `oa check` — Validate PPT against Excel

Cell-by-cell comparison of table values and delta sign verification. Exit code 0 = pass, 1 = mismatches found.

```
oa check <FILE> [OPTIONS]
```

**Options:**

| Flag | Description |
|------|-------------|
| `-e, --excel <PATH>` | Excel to check against (auto-detected if omitted) |
| `--set <KEY=VALUE>` | Override config values (repeatable) |
| `-v, --verbose` | Debug logging |

**What it checks:**
- **Tables**: Every cell in every linked table is compared to its Excel source
- **Transposed tables**: Handles row/col swap correctly
- **_ccst tables**: Applies the same transform (prefix, symbol removal) before comparing
- **Deltas**: Verifies shape sign suffix (_pos/_neg/_none) matches Excel value

**Examples:**

```bash
# Check against specific Excel file
oa check report.pptx -e data.xlsx

# Auto-detect Excel from OLE links
oa check report.pptx

# Use in CI: exit code 1 on mismatch
oa check report.pptx -e data.xlsx || echo "VALIDATION FAILED"
```

**Output on mismatch:**

```
TABLE MISMATCHES (3):
  Slide 5, Object_revenue: Cell (2,1): PPT="45%" vs Expected="52%"
  Slide 5, Object_revenue: Cell (3,1): PPT="30%" vs Expected="25%"
  Slide 8, Object_costs: Cell (1,1): PPT="$1.2M" vs Expected="$1.5M"

CHECK FAILED: 3 table mismatches, 0 delta mismatches (of 150 checked)
```

---

### `oa diff` — Compare two PPTX files

Side-by-side comparison of two presentations. Read-only, no Excel needed.

```
oa diff <A.pptx> <B.pptx> [-v]
```

**What it compares:**
- Shape inventory counts (ntbl_, htmp_, trns_, delt_, _ccst)
- Table cell values for matching shapes
- Chart counts

**Examples:**

```bash
# Compare template vs updated version
oa diff template.pptx updated_report.pptx

# Compare two country reports
oa diff us_report.pptx mx_report.pptx
```

---

### `oa run` — Execute a TOML runfile

Batch processing from a TOML configuration file. Replaces Python runfiles.

```
oa run <RUNFILE.toml> [OPTIONS]
```

**Options:**

| Flag | Description |
|------|-------------|
| `--check` | Validate after each job |
| `--dry-run` | Don't save changes |
| `-v, --verbose` | Debug logging |
| `-q, --quiet` | Errors only |

**TOML Runfile Format:**

```toml
# output/{name}.pptx — {name} is replaced with the job key
default_output = "output/rpm_2024_{name}.pptx"

# Optional: limit which pipeline steps run (default: all)
steps = ["links", "tables", "deltas", "coloring", "charts"]

# Optional: config overrides (same keys as --set)
[config]
ccst.positive_prefix = ""
links.set_manual = true

# Jobs: template path → { job_name = excel_path }
[jobs."templates/region1_template.pptx"]
australia = "data/rpm_tracking_Australia.xlsx"
japan = "data/rpm_tracking_Japan.xlsx"
indonesia = "data/rpm_tracking_Indonesia.xlsx"

[jobs."templates/region2_template.pptx"]
germany = "data/rpm_tracking_Germany.xlsx"
france = "data/rpm_tracking_France.xlsx"

# Per-job output override (use inline table)
[jobs."templates/special_template.pptx"]
usa = "data/usa.xlsx"
canada = { data = "data/canada.xlsx", output = "special/canada_report.pptx" }
```

**Examples:**

```bash
# Run all jobs in the runfile
oa run batch.toml

# Dry run: see what would happen
oa run batch.toml --dry-run

# Run and validate each output
oa run batch.toml --check

# Quiet mode for CI
oa run batch.toml -q
```

---

### `oa config` — Show config keys and defaults

Prints all available `--set` keys with their default values.

```
oa config
```

**Output:**

```
KEY                            DEFAULT
---                            -------
heatmap.color_minimum          #F8696B
heatmap.color_midpoint         #FFEB84
heatmap.color_maximum          #63BE7B
heatmap.dark_font              #000000
heatmap.light_font             #FFFFFF
ccst.positive_color            #33CC33
ccst.negative_color            #ED0590
ccst.neutral_color             #595959
ccst.positive_prefix           +
ccst.symbol_removal            %
delta.template_positive        tmpl_delta_pos
delta.template_negative        tmpl_delta_neg
delta.template_none            tmpl_delta_none
delta.template_slide           1
links.set_manual               true
```

**Config Sections:**

| Section | Keys | Description |
|---------|------|-------------|
| `heatmap.*` | 5 keys | Colors for 3-color scale heatmap tables (htmp_) |
| `ccst.*` | 5 keys | Sign-based color coding (_ccst tables) |
| `delta.*` | 4 keys | Delta indicator template shape names and source slide |
| `links.*` | 1 key | OLE link update behavior |

---

### `oa clean` — Kill zombie Office processes

Finds and kills orphaned POWERPNT.EXE and EXCEL.EXE processes left over from crashes.

```
oa clean [-f]
```

**Options:**

| Flag | Description |
|------|-------------|
| `-f, --force` | Kill without prompting for confirmation |

**Examples:**

```bash
# Interactive: prompts before killing
oa clean

# Force kill (for scripts)
oa clean -f
```

---

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Validation failure (`oa check` found mismatches) |
| 2 | Runtime error (bad arguments, missing files, COM failure) |

---

## Special Shape Naming Conventions

The pipeline identifies shapes by name prefix:

| Prefix | Type | Behavior |
|--------|------|----------|
| `ntbl_` | Normal table | Preserves formatting, only updates cell text |
| `htmp_` | Heatmap table | Recalculates 3-color scale from Excel |
| `trns_` | Transposed table | Swaps rows/columns from Excel range |
| `delt_` | Delta indicator | Arrow shape, swapped based on value sign |
| `_ccst` | Color-coded table | Cells colored by sign (positive/negative/neutral) |
| `tmpl_delta_pos` | Template | Positive delta arrow template on slide 1 |
| `tmpl_delta_neg` | Template | Negative delta arrow template on slide 1 |
| `tmpl_delta_none` | Template | Neutral delta template on slide 1 |

---

## Performance

| Scenario | Time |
|----------|------|
| 30-slide PPTX, 86 OLE, 100 charts | ~10s |
| ZIP pre-relink (186 links) | 0.2s |
| `oa info` inspection | ~3s |
| `oa clean` (no processes) | instant |

Compared to Python `decx` (~40s for the same file), `oa` is approximately **4x faster**.
