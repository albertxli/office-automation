# oa — Office Automation CLI Reference

Windows-only CLI tool that automates Microsoft Office (PowerPoint + Excel) via COM.

## Quick Start

```bash
# Update a single presentation with new Excel data
oa update report.pptx -e data.xlsx

# Batch process 26 countries from a runfile
oa run batch.toml

# Validate all outputs against Excel
oa check batch.toml

# Inspect a PPTX file (read-only)
oa info report.pptx

# Per-slide shape breakdown
oa info -v report.pptx

# Show all config keys and defaults
oa config

# Kill zombie Office processes
oa clean
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

**Pre-pipeline ZIP operations** (run before COM, no PowerPoint needed):

| Operation | Description |
|-----------|-------------|
| ZIP pre-relink | Rewrite OLE/chart paths in PPTX XML (0.1s vs 100s via COM) |
| ZIP chart pre-update | Rewrite chart numCache values directly in XML |

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
oa update report.pptx -e data.xlsx --set ccst.positive_color=#00FF00

# Dry run: see what would happen without saving
oa update report.pptx -e data.xlsx --dry-run

# Update and validate results
oa update report.pptx -e data.xlsx --check

# Auto-detect Excel from OLE links in the PPTX
oa update report.pptx

# Open file dialog to select Excel
oa update report.pptx --pick
```

---

### `oa run` — Execute a TOML runfile

Batch processing from a TOML configuration file. Processes all jobs sequentially using a single shared COM session (avoids 0x80010001 errors from rapid COM create/destroy).

Prints a rich summary table after all jobs complete showing per-job pass/fail, object counts, timing, and totals.

```
oa run <RUNFILE.toml> [OPTIONS]
```

**Options:**

| Flag | Description |
|------|-------------|
| `--check` | Run validation after each job |
| `--dry-run` | Don't save changes |
| `-v, --verbose` | Debug logging |
| `-q, --quiet` | Errors only |

**TOML Runfile Format:**

```toml
# output/{name}.pptx — {name} is replaced with the job key
default_output = "output/{name}.pptx"

# Optional: limit which pipeline steps run (default: all)
steps = ["links", "tables", "deltas", "coloring", "charts"]

# Optional: config overrides (same keys as --set)
[config]
ccst.positive_prefix = ""
links.set_manual = true

# Jobs: template path → { job_name = excel_path }
[jobs."templates/region1_template.pptx"]
australia = "data/tracking_australia.xlsx"
japan = "data/tracking_japan.xlsx"
indonesia = "data/tracking_indonesia.xlsx"

[jobs."templates/region2_template.pptx"]
germany = "data/tracking_germany.xlsx"
france = "data/tracking_france.xlsx"

# Per-job output override (use inline table)
[jobs."templates/special_template.pptx"]
usa = "data/tracking_usa.xlsx"
canada = { data = "data/tracking_canada.xlsx", output = "special/canada_report.pptx" }
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

**Example output:**

```
Runfile: batch.toml (26 jobs)

--- Job 1/26: Argentina ---
  ▸ template.pptx
    ← tracking_argentina.xlsx
  ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌
  • Relink ·······················  411   0.1s
  • Tables ·······················  155   1.2s
  • Deltas ·······················    5   0.4s
  • Coloring ·····················    5   0.1s
  • Charts ·······················  257   0.3s
  ✓ completed · 577 objects · 5.9s

[... 24 more jobs ...]

  ═══════════════════════════════════════
  Job summary

  ✓ Argentina ··················  577 objects   5.9s
  ✓ Australia ··················  577 objects   5.7s
  ✗ Brazil ····················· Excel file not found
  ...

  ✓ all jobs complete · 26/26 files · 15262 objects · 2m 36.1s · avg 5.9s/file
```

---

### `oa check` — Validate PPT against Excel

Cell-by-cell comparison of table values, delta sign verification, and chart data validation. Supports both single PPTX files and batch validation via runfiles.

Exit code 0 = pass, 1 = mismatches found.

```
oa check <FILE> [OPTIONS]
```

**Arguments:**

| Argument | Description |
|----------|-------------|
| `<FILE>` | PPTX file or runfile (`.toml`/`.py`) to validate |

**Options:**

| Flag | Description |
|------|-------------|
| `-e, --excel <PATH>` | Excel to check against (auto-detected if omitted) |
| `--set <KEY=VALUE>` | Override config values (repeatable) |
| `-v, --verbose` | Show per-cell comparison details |

**What it checks:**
- **Tables**: Every cell in every linked table compared to its Excel source
- **Transposed tables**: Handles row/col swap correctly
- **_ccst tables**: Applies the same transform (prefix, symbol removal) before comparing
- **Deltas**: Verifies shape sign suffix (_pos/_neg/_none) matches Excel value
- **Charts**: Link targets, series counts, series values (cached vs Excel)

**Examples:**

```bash
# Check a single file against specific Excel
oa check report.pptx -e data.xlsx

# Auto-detect Excel from OLE links
oa check report.pptx

# Batch check all jobs from a runfile
oa check batch.toml

# Verbose: see every cell comparison
oa check report.pptx -e data.xlsx -v

# Use in CI: exit code 1 on mismatch
oa check report.pptx -e data.xlsx || echo "VALIDATION FAILED"
```

**Example output (single file):**

```
  ▸ report.pptx
    ← data.xlsx
  ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌

  ✓ Tables ·············· 234 checked                    ·     0 mismatches  PASS
  ✓ Deltas ··············   5 checked                    ·     0 mismatches  PASS
  ✓ Charts ·············· 257 checked (546 series)       ·     0 mismatches  PASS

  ✓ check passed · 785 checked · 0 mismatches · 4.3s
```

**Example output (batch via runfile):**

```
  ═══════════════════════════════════════
  Check summary

  ✓ Argentina ··································  785 checked   4.5s
  ✗ Australia ·································· 4 mismatches   4.4s
  ✓ Brazil ·····································  785 checked   4.2s
  ...

  ✗ 1 check failed · 5/6 files · 4740 checked · 27.8s
```

---

### `oa info` — Inspect a PPTX file

Read-only inspection. Shows slide count, OLE links, charts (linked/unlinked), special shapes, and delta templates. With `-v`, adds a per-slide shape breakdown table.

```
oa info <FILE> [-v]
```

**Options:**

| Flag | Description |
|------|-------------|
| `-v, --verbose` | Show per-slide breakdown table |

**Example (normal):**

```bash
oa info template.pptx
```

```
  ▸ template.pptx
  ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌

  File size ···································· 2.7 MB
  Slides ·······································   68

  OLE links ····································  155
    ╰ tracking_data.xlsx ························  155

  Charts ·······································  258
    ╰ Linked ···································  257
    ╰ Unlinked ·································    1

  Special shapes ·······························  165
    ╰ ntbl_ normal tables ······················  122
    ╰ htmp_ heatmap tables ·····················    0
    ╰ trns_ transposed tables ··················   33
    ╰ delt_ delta indicators ···················    5
    ╰ _ccst color-coded ························    5

  Delta templates
    ╰ tmpl_delta_pos ···························    ✓
    ╰ tmpl_delta_neg ···························    ✓
    ╰ tmpl_delta_none ··························    ✓
```

**Example (verbose — per-slide breakdown):**

```bash
oa info -v template.pptx
```

Appends after the normal output:

```
  Per-slide breakdown
  ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌
   slide    ole  chart   ntbl   htmp   trns   delt   ccst  total

       1      ·      ·      ·      ·      ·      ·      ·      ·
       2      5      ·      5      ·      ·      5      5     20
       3      ·     19      ·      ·      ·      ·      ·     19
       4      7      ·      5      ·      2      ·      ·     14
       ...
      68      ·     19      ·      ·      ·      ·      ·     19
  ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌
  68 slides · 42 active · 26 empty
```

Columns: `slide` (slide number), `ole` (OLE objects), `chart` (linked charts only), `ntbl` (normal tables), `htmp` (heatmap tables), `trns` (transposed tables), `delt` (delta indicators), `ccst` (color-coded shapes), `total` (sum). Zero values shown as `·`. All slides shown including empty ones.

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

### `oa config` — Show config keys and defaults

Prints all available `--set` keys with their default values.

```
oa config
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

Finds and kills orphaned POWERPNT.EXE and EXCEL.EXE processes left over from crashes. Shows found processes with PIDs, prompts for confirmation before killing.

```
oa clean [-f]
```

**Options:**

| Flag | Description |
|------|-------------|
| `-f, --force` | Kill without prompting for confirmation |

**Examples:**

```bash
# Interactive: lists processes and prompts before killing
oa clean

# Force kill (for scripts)
oa clean -f
```

**Example output:**

```
  Found 2 Office processes

  EXCEL.EXE ························ PID 95448
  POWERPNT.EXE ····················· PID 99640

  Kill all? [y/N] y

  ✓ Killed EXCEL.EXE ················· PID 95448
  ✓ Killed POWERPNT.EXE ·············· PID 99640

  ✓ cleaned · 2 processes killed
```

When no processes found:

```
  Found 0 Office processes

  ✓ No Office processes found
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

The pipeline identifies shapes by name prefix/suffix:

| Prefix/Suffix | Type | Behavior |
|---------------|------|----------|
| `ntbl_` | Normal table | Preserves formatting, only updates cell text |
| `htmp_` | Heatmap table | Recalculates 3-color scale from Excel |
| `trns_` | Transposed table | Swaps rows/columns from Excel range |
| `delt_` | Delta indicator | Arrow shape, swapped based on value sign |
| `_ccst` | Color-coded table | Cells colored by sign (positive/negative/neutral) |
| `tmpl_delta_pos` | Template | Positive delta arrow template on slide 1 |
| `tmpl_delta_neg` | Template | Negative delta arrow template on slide 1 |
| `tmpl_delta_none` | Template | Neutral delta template on slide 1 |

**Shape-OLE matching:** Table names like `ntbl_Object 1_ccst` are matched to OLE shapes like `Object 1` using word-boundary token matching (the `ntbl_` prefix and `_ccst` suffix are stripped during matching).

**Delta empty data handling:** When the Excel cell for a delta indicator is empty/missing, the delta shape is set to `_none` (neutral indicator) rather than being skipped.

---

## Performance

| Scenario | Time |
|----------|------|
| Single 68-slide PPTX (155 OLE, 257 charts) | ~6s |
| Batch 26 files via `oa run` | ~2m 36s |
| ZIP pre-relink (411 links) | 0.1s |
| ZIP chart pre-update (257 charts) | 0.3s |
| `oa info` inspection | ~3s |
| `oa check` single file | ~4s |
| `oa check` batch (6 files via runfile) | ~28s |
| `oa clean` (no processes) | instant |

**Key optimization:** COM session reuse across batch jobs saves ~28s on 26 jobs by avoiding rapid COM create/destroy (GOTCHA #39).

**Limitation:** PowerPoint is a single-instance COM server (GOTCHA #40). Multiple threads all share one POWERPNT.EXE process, so multi-threaded parallelism provides no speedup for PowerPoint operations.
