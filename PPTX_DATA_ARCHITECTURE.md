# PPTX Data Architecture Reference

How data flows through OLE objects and charts in PowerPoint files, and how `oa` interacts with each layer.

## PPTX File Structure (ZIP Archive)

A `.pptx` file is a ZIP containing:

```
docProps/
  core.xml              ← Author, modified date, revision count
  app.xml               ← Slide count, format, fonts, template, word count
ppt/
  presentation.xml      ← Slide ordering, dimensions
  slides/
    slide1.xml          ← Shape positions, graphicFrames, OLE refs
    _rels/
      slide1.xml.rels   ← OLE link paths (file:///...) + chart rId mappings
  charts/
    chart1.xml          ← Series data (<c:numCache>), range formulas (<c:f>)
    _rels/
      chart1.xml.rels   ← Chart external link path (bare path or file:///)
  embeddings/           ← Embedded OLE data (Excel sheets as binary)
  media/                ← Images (EMF snapshots of OLE objects)
```

## OLE Objects — Data Layers

| Layer | Location | What it tells you | Access cost |
|-------|----------|-------------------|-------------|
| **Link path** | `slides/_rels/slideN.xml.rels` → `Target` attribute | Which Excel file + sheet + range this OLE points to | ZIP read (fast) |
| **Embedded image** | `ppt/media/imageN.emf` | What the user SEES (frozen snapshot) | ZIP read (fast) |
| **COM SourceFullName** | `LinkFormat.SourceFullName` via IDispatch | Same path as .rels, but as PowerPoint sees it at runtime | COM call (slow) |
| **COM ProgID** | `OLEFormat.ProgID` via IDispatch | Object type (e.g., `Excel.Sheet.12`) | COM call (slow) |
| **Excel cell values** | `Range.Text` / `Range.Value2` via Excel COM | The actual source-of-truth data | Excel COM (slow) |

**Key insight**: The OLE object in the slide is a **frozen EMF image** with a link pointer. The image doesn't update until PowerPoint refreshes it. What `oa update` does is write the text directly into the PPT table cells (via COM), not refresh the OLE image.

## Charts — Data Layers

| Layer | Location | What it tells you | Access cost |
|-------|----------|-------------------|-------------|
| **Cached values** | `ppt/charts/chartN.xml` → `<c:numCache>/<c:pt>/<c:v>` | Data points PowerPoint **renders and displays** | ZIP read (fast) |
| **Range formula** | `ppt/charts/chartN.xml` → `<c:val>/<c:numRef>/<c:f>` | Which Excel cells this series references | ZIP read (fast) |
| **Format code** | `ppt/charts/chartN.xml` → `<c:formatCode>` | How values are formatted (e.g., `0%`, `General`) | ZIP read (fast) |
| **External link** | `ppt/charts/_rels/chartN.xml.rels` → `Target` | Excel file the chart links to | ZIP read (fast) |
| **COM Series.Values** | `Series.Values` via IDispatch SAFEARRAY | What COM returns when you query chart data | COM call (slow) |
| **COM IsLinked** | `Chart.ChartData.IsLinked` via IDispatch | Whether chart has external data source | COM call (slow) |
| **Excel cell values** | `Range.Value2` via Excel COM SAFEARRAY | The **source of truth** for what data should be | Excel COM (moderate) |

**Key insight**: PowerPoint renders charts from `<c:numCache>` — that's what the user sees on screen. If the cache is updated, the chart displays correctly even without calling `LinkFormat.Update()`.

## The Verification Chain

```
Excel Range.Value2  →  must match  →  ZIP <c:numCache>  →  what user sees
     (source)                          (cache)               (rendered)
```

- `oa update` writes: Excel → ZIP cache (via chart_data.rs)
- `oa check` verifies: ZIP cache ↔ Excel values (via check.rs)
- Old approach (COM): Excel → COM Update() → internal workbook → cache (slow, unreliable)

## Presentation Metadata (ZIP XML, no COM)

### `docProps/core.xml`

| Field | XML element | Example |
|-------|-------------|---------|
| Creator/Author | `<dc:creator>` | `Albert Li` |
| Last modified by | `<cp:lastModifiedBy>` | `Albert Li` |
| Created date | `<dcterms:created>` | `2025-03-10T15:33:08Z` |
| Modified date | `<dcterms:modified>` | `2026-03-24T02:31:00Z` |
| Revision count | `<cp:revision>` | `574` |

### `docProps/app.xml`

| Field | XML element | Example |
|-------|-------------|---------|
| Presentation format | `<PresentationFormat>` | `Widescreen` |
| Template name | `<Template>` | `Povaddo Theme` |
| Slide count | `<Slides>` | `71` |
| Notes count | `<Notes>` | `11` |
| Hidden slides | `<HiddenSlides>` | `0` |
| Word count | `<Words>` | `9180` |
| Fonts used | `<HeadingPairs>` + `<TitlesOfParts>` | 13 fonts |
| Office version | `<AppVersion>` | `16.0000` |
| Total editing time | `<TotalTime>` | `52611` (minutes) |

## How `oa` Uses Each Layer

### `oa update`

1. **ZIP relinker** — rewrites `slides/_rels/` and `charts/_rels/` link paths
2. **ZIP chart data** — reads Excel via `Range.Value2` SAFEARRAY, writes `<c:numCache>/<c:pt>` values
3. **COM links** — sets `AutoUpdate=Manual` on each OLE/chart
4. **COM tables** — reads Excel `.Text` per cell, writes to PPT table cells
5. **COM deltas** — reads Excel values, swaps delta indicator shapes
6. **COM coloring** — reads cell text, applies sign-based color coding

### `oa check`

1. **ZIP chart cache** — reads `<c:numCache>` values (PPT side, no COM)
2. **ZIP chart refs** — reads `<c:f>` range formulas
3. **Excel Range.Value2** — reads expected values via SAFEARRAY (Excel COM)
4. **COM tables** — reads PPT `.Text` and Excel `.Text` per cell for comparison
5. **COM link check** — reads `SourceFullName` to verify link targets

### `oa info`

1. **ZIP metadata** — file size from filesystem
2. **COM inventory** — shape discovery (OLE, charts, tables, deltas)
3. **ZIP chart cache** — empty cache detection

## Known Edge Cases

| # | Issue | Gotcha | Impact |
|---|-------|--------|--------|
| 20 | Non-contiguous chart ranges `(A1:A5,C1:C5)` | Split on commas, read each sub-range | Update + check |
| 23 | Chart XML: filter to `<c:val>` only, not `<c:cat>` | Ignore category axis data | Update + check |
| 25 | `IsLinked=True` but `SourceFullName="NULL"` | Skip broken links gracefully | Update + check |
| 34 | ZIP rename fails (SynologyDrive sync lock) | Graceful fallback to COM pipeline | Update |
| 35 | Chart .rels use bare paths (no `file:///` prefix) | Relinker handles both formats | Update |
| 36 | Empty `<c:numCache>` (ptCount but no pt elements) | Inject pt elements from Excel data | Update + check |

## Future Improvement Opportunities

1. **`oa info -v`** — per-slide breakdown (OLE count, chart count, table count per slide)
2. **`docProps` metadata in `oa info`** — last modified by, modified date, revision count
3. **ZIP-only table text update** — write `<a:t>` elements in slide XML directly, bypass COM table writes
4. **Parallel batch processing** — rayon with per-thread STA COM for `oa run` multi-file
5. **Chart category axis update** — currently only `<c:val>` is updated, `<c:cat>` labels are left as-is
6. **Embedded chart workbook sync** — update the internal Excel workbook inside the chart (not just cache)
