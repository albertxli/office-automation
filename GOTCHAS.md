# COM Gotchas

Hard-won lessons from the Python `decx` development. Every one of these cost debugging time.

## Shape & OLE Constants

### #1 Security Dialog
PowerPoint shows a "Microsoft PowerPoint Security Notice" dialog when opening files with OLE links. This blocks COM automation. **Solution:** Background thread using `FindWindowW` + `PostMessageW(WM_CLOSE)` to auto-dismiss.

### #2 Wrong OLE Constant
`msoLinkedOLEObject = 10`, NOT 7. The value 7 is `msoEmbeddedOLEObject`. Always verify constants from the Office type library.

### #11 Wrong Delta Constant
`ppUpdateOptionManual = 1`, `ppUpdateOptionAutomatic = 2`. Easy to confuse.

## COM Lifecycle

### #3 Zombie Cleanup
After `Application.Quit()`, COM objects may still hold RPC connections. Need garbage collection + brief delay, or force-kill zombie processes if cleanup fails.

### #4 DispatchEx vs Dispatch
Always use `DispatchEx` (or `CoCreateInstance` with `CLSCTX_LOCAL_SERVER` in Rust) for process isolation. `Dispatch` can reuse existing instances, causing interference.

### #9 Zombie Processes
After crashes, `POWERPNT.EXE` and `EXCEL.EXE` zombies persist. The `oa clean` command handles this.

### #21 Excel.Quit() Hangs (60 seconds!)
If `IDispatch` references from inventory/shapes still exist when `Application.Quit()` is called, the process hangs for ~60 seconds waiting for RPC disconnection. **Solution:** Explicitly drop all shape/inventory references BEFORE calling Quit. Drop order: inventory → presentation → Excel quit → PPT quit.

## Excel Operations

### #5 Excel UpdateLinks
Pass `UpdateLinks=0` when opening workbooks to prevent auto-refresh of external links.

### #6 Excel Calculation Mode
Set to manual (`xlCalculationManual = -4135`) during processing to prevent recalculation on every cell read. Restore to automatic (`-4105`) after.

### #16 Range.Value2 Loses Formatting
Use `.Text` for display strings (preserves number formatting). `.Value2` returns raw numeric values that lose currency symbols, percentages, etc.

## Link Operations

### #12 Bulk UpdateLinks is SLOWER
`Presentation.UpdateLinks()` is actually SLOWER than per-shape `LinkFormat.Update()`. Use per-shape updates.

### #15 LinkFormat.Update() Performance
Extremely slow for remote/network files (~4 seconds per shape). **Skip it entirely** for OLE links — the ZIP pre-relink handles path changes. Only call `Update()` for charts (which need data refresh).

### #22 ZIP Pre-Relink
Rewriting OLE/chart link paths directly in the PPTX XML (which is a ZIP archive) before opening via COM eliminates the per-link COM overhead. Performance: 0.12s for 186 links vs ~90-100s via COM.

### #25 Broken Chart Links
Some charts have `IsLinked=True` but `SourceFullName="NULL"`. This is acceptable — skip these gracefully rather than erroring.

## Shape Discovery

### #13 Group Delta Shapes
When scanning grouped shapes, check the **group's name** for `delt_` prefix BEFORE recursing into its children. The group itself is the delta shape, not its children.

### #17 Inventory Key by Slide
Shape names can be duplicated across slides (e.g., two slides both have an OLE named `Table_Revenue`). Key the inventory by `(slide_index, shape_name)` tuple, not just shape name.

## Table Operations

### #7 Float Precision in Contrast Color
The luminance formula `0.299*R + 0.587*G + 0.114*B` with threshold `< 128` can produce floating-point edge cases (e.g., `0.299 * 128 = 127.9999...` in some implementations). Use `<= 127.5` or careful rounding.

## Chart Operations

### #18 Series.Formula Inaccessible
On linked charts, `Series.Formula` is inaccessible (raises error). Use `.Values` property to read data and parse the chart XML for range references instead.

### #19 Chart Mapping via Slide/Position
Charts in PPTX XML don't use flat indexing. Map charts to their slide and position, not by XML filename order.

### #20 Non-Contiguous Chart Ranges
Chart data ranges can be non-contiguous (comma-separated in the formula). Split on commas and read each range part separately.

### #23 Chart XML: Filter to val/ Only
When collecting `numRef` elements from chart XML, only collect those inside `<c:val>` elements (value axis). Ignore those in `<c:cat>` (category axis) to avoid double-counting.

### #24 Unlinked Charts in XML
Filter `.rels` entries to external references only. Internal chart data (embedded) should not be treated as linked charts.

## Miscellaneous

### #8 VBA Reference
The authoritative reference is `RUN ALL_Table+Chart_v11.bas`, not the old Jupyter notebook.

### #10 ZIP Corruption
Be careful with tools that modify PPTX ZIP structure — verify the output is still valid.

### #33 Presentations.Open Untitled=True Breaks Save (Rust-specific)
Opening with `Untitled=True` (-1) makes PowerPoint treat the file as a new unnamed document. `Save()` then silently does nothing — all changes are lost. **Always use `Untitled=False` (0)** for files you intend to modify and save. This bug caused table cell writes and chart updates to appear successful but produce no changes in the output file.

### #32 Skip COM SourceFullName — Use ZIP Pre-Relink Only (Rust-specific)
`LinkFormat.SourceFullName = "..."` triggers PowerPoint to resolve/validate the link target: file I/O, network check, sheet/range validation. This costs **0.5s per shape** (86 shapes = 42.7s). ZIP pre-relink already rewrites all paths in 0.2s directly in the PPTX XML. When PowerPoint opens the file, it reads the ALREADY-CORRECT paths. **Never set SourceFullName via COM** — only use COM to set `AutoUpdate=Manual` (~0.01s/shape). This single optimization took pipeline time from 54s to 9.8s.

### #31 DISPIDs Are Per-COM-Class, Not Global (Rust-specific)
DISPIDs are unique per COM class, NOT globally unique. "Item" on a Slides collection has a DIFFERENT DISPID than "Item" on a Shapes collection. A global DISPID cache causes "Member not found" errors. Use per-instance caching with Rc<RefCell<HashMap>> shared across clones of the same object.

### #30 Dry-Run Must Skip ZIP Pre-Relink (Rust-specific)
ZIP pre-relink modifies the PPTX file on disk. `--dry-run` must skip this step, otherwise the file is permanently modified even though the COM pipeline doesn't save. This is especially dangerous in in-place mode (no `-o`).

### #29 Use CLI Excel Path, Not SourceFullName Path (Rust-specific)
The `SourceFullName` in OLE links contains the ORIGINAL Excel path (from whoever created the template), which likely doesn't exist on the current machine. Always use the Excel path from the command line (`-e` flag) for opening workbooks. Only use SourceFullName for extracting the sheet name and cell range.

### #28 Excel.Calculation Before Workbooks Open (Rust-specific)
Setting `Excel.Application.Calculation = xlCalculationManual` on a fresh Excel instance with NO open workbooks throws `DISP_E_EXCEPTION (0x80020009)`. **Defer the Calculation mode change until after opening the first workbook.** Python's pywin32 may have masked this by silently swallowing the exception.

### #26 Collection Indexer is call(), not get_with() (Rust-specific)
COM collections (Slides, Shapes, Worksheets, etc.) expose their `Item` indexer as a method, not a parameterized property. In Rust: use `dispatch.call("Item", &[index])`, NOT `dispatch.get_with("Item", &[index])`. The `get_with` silently fails with "Member not found" (DISP_E_MEMBERNOTFOUND).

### #27 Windows \\\\?\ UNC Prefix (Rust-specific)
`std::path::Path::canonicalize()` on Windows produces `\\?\C:\...` extended-length paths. Office COM does NOT understand these — `Workbooks.Open()` will fail with `DISP_E_EXCEPTION`. Always strip the `\\?\` prefix before passing paths to COM.

### #34 ZIP Pre-Relink "Access is denied" (os error 5) (Rust-specific)
`std::fs::rename()` of `.pptx.tmp` over the original PPTX fails with os error 5 when another process has a file handle — typically SynologyDrive sync agent, Windows Search indexer, antivirus, or Explorer preview pane. Common on first run when the PPTX was last modified on another machine (sync agent locks on detecting a new file). On second run the file is stable and rename succeeds. **The fallback is correct by design:** `update.rs` catches the error, prints a warning, and continues with the original PPTX. COM pipeline updates links normally (slower but correct). ZIP pre-relink is a performance optimization, not required for correctness.

### #36 Empty Chart numCache (ptCount But No pt Elements) (Rust-specific)
Some charts have `<c:numCache>` with `<c:ptCount val="N"/>` but **zero `<c:pt>` data point elements**. The cache structure exists but contains no actual values. This happens when charts are duplicated in the template without their cache being populated, or when the template was created by a tool that didn't fill the cache.

**Impact on oa update:** The ZIP chart data pre-update rewrites existing `<c:pt>/<c:v>` text but can't replace what doesn't exist. **Fix:** Detect the empty-cache case (no `<c:pt>` seen between `<c:numCache>` start and end) and **inject** new `<c:pt idx="N"><c:v>VALUE</c:v></c:pt>` elements from Excel data before writing `</c:numCache>`.

**Impact on oa check:** Reading cached values from ZIP returns an empty Vec for these series. Comparing `[]` vs `[0.74]` from Excel reports a false mismatch. **Fix:** Skip comparison when cached values are empty (the chart data is unverifiable from the ZIP cache alone).

### #37 Partial Chart numCache (ptCount > Number of pt Elements) (Rust-specific)
Some charts have `<c:numCache>` with `<c:ptCount val="3"/>` but only 2 `<c:pt>` elements — a partial cache. This happens with multi-series/multi-column charts where the latest wave or year column hasn't been populated yet (e.g., a 3-column chart for 2023/2024/2025 where 2025 data isn't collected yet).

**Impact on oa update:** The ZIP chart data pre-update replaces existing `<c:pt>` values but didn't create the missing trailing ones. **Fix:** Track `max_pt_idx_seen` during cache traversal. On `</c:numCache>`, inject any remaining values from `max_pt_idx_seen` to `vals.len()`. This scales dynamically for any number of columns/series.

**Impact on oa check:** The ZIP cache has fewer values than Excel, causing `1/N differ` mismatches on the missing column. **Fix:** Same injection in oa update ensures the cache is complete.

### #35 Chart .rels Use Bare Paths (No file:/// Prefix) (Rust-specific)
OLE link `.rels` in `slides/_rels/` use `file:///C:/path/to/file.xlsx` format. But chart `.rels` in `charts/_rels/` use bare paths: `C:/path/to/file.xlsx` (no `file:///` prefix). The relinker must handle both formats — match and rewrite bare paths as well as `file:///` URIs. Preserve the original format when writing back.

### #14 Rich Unicode
Terminal output can use Unicode freely. Earlier cp1252 limitation was a test environment issue only.
