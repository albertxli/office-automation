# office-automation (`oa`)

Windows-only CLI tool that automates Microsoft Office (Word, Excel, PowerPoint) via the Windows COM API. Rust rewrite of Python `decx`.

## Build

```bash
cargo build --release
cargo test                              # Unit tests (no Office needed)
OA_INTEGRATION=1 cargo test             # All tests (needs Office installed)
cargo clippy -- -D warnings
```

## Run

```bash
cargo run --release -- --help           # Show all commands
cargo run --release -- config           # Show config keys
cargo run --release -- info FILE.pptx   # Inspect a PPTX
```

## Architecture

```
COM wrapper (com/)        → IDispatch wrapper with DISPID caching, VARIANT newtype, RAII guards
Typed Office (office/)    → PowerPointApp, ExcelApp, Presentation, Slide, Shape, Range
Pipeline (pipeline/)      → linker, table_updater, delta_updater, color_coder, chart_updater
ZIP ops (zip_ops/)        → PPTX ZIP manipulation without COM (pre-relink, detection)
Shapes (shapes/)          → Inventory, matcher, formatting
Commands (commands/)      → CLI command implementations
Utils (utils/)            → Color, cell refs, link parsing, process management
```

## Key Conventions

- **Shape types**: `ClassifiedShape` enum with `enum_dispatch` — zero-cost dispatch, not string prefix matching
- **Arena allocation**: `bumpalo` for per-presentation shape inventories
- **String interning**: `lasso` for shape name matching (O(1) comparison)
- **DISPID caching**: `HashMap` in `Dispatch` wrapper to skip repeated `GetIDsOfNames` calls
- **Parallel batch**: `rayon` with per-thread STA COM apartments for independent Office instances
- **Drop ordering**: inventory → presentation → Excel quit → PPT quit → COM guard (prevents 60s hang)
- **ZIP pre-relink**: Rewrite OLE/chart paths directly in PPTX XML before COM opens file (0.12s vs 100s)

## COM Gotchas

See `GOTCHAS.md` for all 25 documented issues from Python development.

**Critical ones to always check:**
- #2: `msoLinkedOLEObject = 10` (not 7)
- #13: Check group name for `delt_` BEFORE recursing into children
- #15: Skip `LinkFormat.Update()` for OLE (extremely slow ~4s/shape)
- #16: Use `.Text` not `.Value2` for display strings from Excel
- #17: Inventory keyed by `(slide_idx, ole_name)` for duplicate names across slides
- #21: Delete all Dispatch refs before calling `Application.Quit()`
- #25: `IsLinked=True` but `SourceFullName="NULL"` is acceptable for broken chart links

## Test Data

- `quick_test_files/` — Template PPTX + 3 country Excel/PPTX pairs + chart validation files
- `excel_test_data/` — Full 27-country Excel dataset
- `example_runfile.py` — Python runfile reference (4 regions, 26 markets)

## Python Reference

Original Python source: `C:\Users\lipov\SynologyDrive\ppt-automation\src\decx\`
