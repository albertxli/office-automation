# Plan: Rust-First `oa` CLI — Office Automation via COM

## Context

Migrating Python `decx` (reference at `C:\Users\lipov\SynologyDrive\ppt-automation`) to a **Rust-native** CLI. This is NOT a line-by-line port — we use the Python as behavioral reference only and design from scratch for Rust's strengths: zero-cost abstractions, fearless concurrency, compile-time type safety, and direct Windows API access.

**Goal:** Feature parity in behavior, but Rust-idiomatic internals that maximize speed. Python processes a 30-slide/86-OLE/100-chart deck in ~40s — we target <15s single file, with true parallel batch processing.

**Decisions made:**
- TOML-only runfiles (clean break from Python runfiles)
- Include `--pick` (native file dialog via `rfd`) in v0.1
- Build strictly sequential: Phase 0 → Phase 1 → Phase 2 → ...

## Test Data (already in project)

- `quick_test_files/` — Template PPTX + 3 country Excel/PPTX pairs + chart validation files
- `excel_test_data/` — Full 27-country Excel dataset
- `example_runfile.py` — Production runfile example (4 regions, 26 markets) — reference for TOML migration
- `test_runfile.py` — Quick 3-market test runfile

---

## CLI Design (Improved from Python)

Changes from Python `decx`:
- **Remove `steps` command** → content folded into `oa update --help`
- **`--pair PPT=XLSX`** (not `:`) → avoids Windows drive-letter colon ambiguity
- **`--steps links,tables`** replaces `--only STEP` (repeatable) → comma-separated, cleaner
- **`--skip links,charts`** replaces individual `--skip-links`/`--skip-charts` flags → scales to new steps
- **Add `--dry-run`** on `update` and `run` → safety net for in-place modifications
- **Add `--check`** on `update` → parity with `run --check`
- **Add `-q/--quiet`** → suppress output except errors (for scripts/CI)
- **Standardized exit codes:** 0=success, 1=validation failure, 2=runtime error

```
oa update <FILES...>                   Main pipeline
  -e, --excel <PATH>                   Excel data file (auto-detected if omitted)
  -p, --pick                           Native file dialog to select Excel
  --pair <PPT=XLSX>                    Explicit pair (repeatable)
  -o, --output <PATH>                  Output file or directory
  --steps <STEP,...>                   Run only these steps (comma-sep)
  --skip <STEP,...>                    Skip these steps (comma-sep, mutex with --steps)
  --set <KEY=VALUE>                    Config override (repeatable)
  --check                             Validate after processing
  --dry-run                           Show what would happen, don't save
  -v, --verbose                        Debug logging
  -q, --quiet                          Errors only

oa run <RUNFILE.toml>                  Batch from TOML runfile
  --check                             Validate after each job
  --dry-run
  -v, --verbose
  -q, --quiet

oa check <FILE> [-e EXCEL]             Validate PPT vs Excel
  --set <KEY=VALUE>
  -v, --verbose

oa diff <A.pptx> <B.pptx>             Compare two PPTX files
  -v, --verbose

oa info <FILE>                         Inspect PPTX (read-only)

oa clean [-f]                          Kill zombie Office processes

oa config                              Show all --set keys and defaults
```

---

## Rust-First Performance Strategy

| Technique | Where | Why |
|-----------|-------|-----|
| **Enum dispatch** (`enum_dispatch`) | Shape types | 10x faster than dyn Trait, zero vtable vs Python string prefix |
| **Rayon** | Batch processing, ZIP/XML | True parallelism; Python's parallel was broken |
| **Per-thread STA COM** | Batch `--pair` / `run` | Independent Office instances per thread |
| **Arena allocation** (`bumpalo`) | Shape inventory per presentation | Contiguous memory, one bulk dealloc |
| **String interning** (`lasso`) | Shape name matching | O(1) symbol comparison vs substring scans |
| **Streaming XML** (`quick-xml`) | ZIP pre-relink, chart parsing | Single-pass transform, no DOM |
| **DISPID caching** | Dispatch wrapper | Cache COM method IDs, skip repeated lookups |
| **Smart drop ordering** | COM cleanup | Prevents 60s hang (gotcha #21) |

---

## Dependencies (all latest, no legacy)

```toml
[package]
name = "office-automation"
version = "0.1.0"
edition = "2024"

[[bin]]
name = "oa"
path = "src/main.rs"

[dependencies]
clap = { version = "4.6", features = ["derive"] }
windows = { version = "0.62", features = [
    "Win32_Foundation",
    "Win32_System_Com",
    "Win32_System_Ole",
    "Win32_System_Variant",
    "Win32_System_Threading",
    "Win32_UI_WindowsAndMessaging",
] }
zip = "8.2"
quick-xml = "0.39"
serde = { version = "1.0", features = ["derive"] }
toml = "1.0"
thiserror = "2.0"
anyhow = "1.0"
tracing = "0.1"
tracing-subscriber = "0.3"
indicatif = "0.18"
colored = "3.1"
glob = "0.3"
sysinfo = "0.38"
rfd = "0.17"
regex = "1"
rayon = "1.10"
bumpalo = { version = "3.17", features = ["collections"] }
lasso = "0.7"
enum_dispatch = "0.3"

[dev-dependencies]
tempfile = "3.25"
assert_cmd = "2.1"
predicates = "3.1"
```

---

## Project Structure

```
office-automation/
  Cargo.toml
  CLAUDE.md                         # Build commands, conventions, gotchas reference
  GOTCHAS.md                        # All 25 COM gotchas from Python experience
  plan.md                           # This file

  src/
    main.rs                         # Entry: clap parse → command dispatch
    lib.rs                          # Crate root re-exports
    cli.rs                          # Clap derive: Cli, Commands, all flags
    config.rs                       # Config struct + Default + --set overrides
    error.rs                        # OaError enum (thiserror)

    com/
      mod.rs
      variant.rs                    # Variant newtype with From/Into conversions
      dispatch.rs                   # Dispatch wrapper with DISPID caching
      session.rs                    # OfficeSession: STA, app creation, security dialog
      cleanup.rs                    # RAII guards, drop ordering, zombie prevention

    office/
      mod.rs
      app.rs                        # PowerPointApp + ExcelApp wrappers
      presentation.rs               # Presentation, Slide iteration
      shapes.rs                     # ClassifiedShape enum via enum_dispatch
      excel_data.rs                 # Workbook, Worksheet, Range (batch-read optimized)
      constants.rs                  # MsoShapeType, PpUpdateOption, XlCalculation

    pipeline/
      mod.rs                        # Step sequencing, progress, batch orchestration
      linker.rs                     # OLE link re-pointing
      table_updater.rs              # Table population from Excel
      delta_updater.rs              # Delta indicator swapping
      color_coder.rs                # _ccst sign-based coloring
      chart_updater.rs              # Chart link + data refresh

    zip_ops/
      mod.rs                        # ZIP-level PPTX operations (no COM)
      relinker.rs                   # Rewrite link paths in .rels XML
      detector.rs                   # Auto-detect linked Excel
      xml_stream.rs                 # Streaming XML transform helpers

    shapes/
      mod.rs
      inventory.rs                  # Arena-backed SlideInventory + interned names
      matcher.rs                    # Shape name classification (lasso)
      formatting.rs                 # CellFormatting/TableFormatting

    commands/
      mod.rs
      update.rs                     # oa update (single + parallel batch)
      info.rs                       # oa info
      config_cmd.rs                 # oa config
      run.rs                        # oa run (TOML runfiles + rayon)
      check.rs                      # oa check
      diff.rs                       # oa diff
      clean.rs                      # oa clean

    utils/
      mod.rs
      color.rs                      # hex↔BGR, contrast font color
      cell_ref.rs                   # R1C1 ↔ A1 conversion
      link_parser.rs                # Parse OLE SourceFullName
      process.rs                    # sysinfo zombie kill
      file_picker.rs                # rfd native dialog
```

---

## Build Phases (TDD — tests before every feature)

### Phase 0: Skeleton + CLAUDE.md ⬜
**Output:** Compiles, `oa --help` works, all stubs in place.

- [ ] `cargo init`, full `Cargo.toml`
- [ ] All module files with `mod` declarations
- [ ] `cli.rs`: Complete clap derive — 7 subcommands + all flags
- [ ] `config.rs`: `Config` with `Default` matching Python's `DEFAULT_CONFIG`
- [ ] `error.rs`: `OaError` enum
- [ ] `CLAUDE.md` + `GOTCHAS.md`
- [ ] **Test:** `cargo build` passes, `oa --help` prints all commands

### Phase 1: COM Foundation ⬜ (highest risk)
**Output:** Can create Office COM instances, read/write properties.

- [ ] `com/variant.rs` + **tests**: round-trip all types, null/empty
- [ ] `com/dispatch.rs` + **tests**: get/put/call, DISPID caching, nav chaining
- [ ] `com/cleanup.rs` + **tests**: RAII guard drop ordering
- [ ] `com/session.rs` + **tests**: STA init, security dialog dismisser
- [ ] **Integration test `tests/com_smoke.rs`**: Create Excel → open workbook → read cell → close → quit → no zombies

### Phase 2: Office Wrappers + Utils + `oa info` ⬜
**Output:** Typed Office API, `oa info` works end-to-end.

- [ ] `office/constants.rs`: All COM constant enums
- [ ] `office/app.rs` + `presentation.rs` + `excel_data.rs`: Typed wrappers
- [ ] `office/shapes.rs`: `ClassifiedShape` enum with `enum_dispatch`
- [ ] `utils/color.rs` + **tests**: hex↔BGR, contrast color (incl. float precision gotcha #7)
- [ ] `utils/cell_ref.rs` + **tests**: R1C1↔A1 (single cell, multi-col >26, absolute refs)
- [ ] `utils/link_parser.rs` + **tests**: all SourceFullName formats
- [ ] `shapes/matcher.rs` + **tests**: classify all prefixes, token matching
- [ ] `shapes/inventory.rs` + **tests**: arena alloc, (slide_idx, name) keying, group recursion, priority
- [ ] `commands/info.rs` + **integration test**: `oa info quick_test_files/test_template.pptx`

### Phase 3: ZIP Operations ⬜ (pure Rust, no COM)
**Output:** Fast PPTX relinking.

- [ ] `zip_ops/xml_stream.rs` + **tests**: XML identity round-trip, attribute rewriting
- [ ] `zip_ops/relinker.rs` + **tests**: rewrite known link paths, verify XML
- [ ] `zip_ops/detector.rs` + **tests**: detect Excel from test PPTX
- [ ] **Integration test**: relink `quick_test_files/test_template.pptx`, verify via `oa info`

### Phase 4: Update Pipeline ⬜ (core value)
**Output:** `oa update` works end-to-end.

- [ ] `pipeline/linker.rs` + **tests**: SourceFullName set, AutoUpdate=1, skip Update()
- [ ] `shapes/formatting.rs` + **tests**: extract/apply round-trip, minimal extraction
- [ ] `pipeline/table_updater.rs` + **tests**: .Text reads, ntbl_/htmp_/trns_ paths, heatmap
- [ ] `pipeline/delta_updater.rs` + **tests**: sign determination, two-pass, template copy
- [ ] `pipeline/color_coder.rs` + **tests**: numeric detection, coloring, prefix/symbol
- [ ] `pipeline/chart_updater.rs` + **tests**: link update, broken links (gotcha #25)
- [ ] `pipeline/mod.rs` + **tests**: step sequencing, --steps/--skip, progress
- [ ] `commands/update.rs`: single + batch (rayon), --dry-run, --check
- [ ] **Integration test**: full pipeline on test data

### Phase 5: Remaining Commands ⬜
**Output:** Full feature parity.

- [ ] `commands/check.rs` + **tests**: cell comparison, delta signs, chart series, exit codes
- [ ] `commands/diff.rs` + **tests**: table/delta/chart diffs, shapes only in A/B
- [ ] `commands/clean.rs` + **tests**: process detection, --force
- [ ] `commands/config_cmd.rs` + **tests**: all keys, format
- [ ] `commands/run.rs` + **tests**: TOML parsing, validation, path resolution, batch rayon
- [ ] `utils/file_picker.rs`: rfd native dialog
- [ ] `utils/process.rs`: zombie helpers

### Phase 6: Polish + Final Integration ⬜
- [ ] --verbose, --quiet, --set, --output, Windows paths
- [ ] Colored output, progress spinners
- [ ] End-to-end CLI tests (assert_cmd)
- [ ] Performance benchmarks vs Python
- [ ] Finalize CLAUDE.md + GOTCHAS.md

---

## Performance Targets

| Scenario | Python | Rust Target |
|----------|--------|-------------|
| Single file (30 slides, 86 OLE, 100 charts) | ~40s | <15s |
| Batch 3 files (sequential) | ~64s | <25s |
| Batch 3 files (parallel rayon) | N/A (broken) | <15s |
| ZIP pre-relink 186 links | 0.12s | <0.05s |
