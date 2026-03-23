# Local Testing Guide

## Prerequisites

- Windows with Microsoft Office (PowerPoint + Excel) installed
- Rust toolchain (`rustup`, `cargo`)

## Install

```bash
# Build and install globally (oa.exe goes to ~/.cargo/bin/)
cargo install --path .

# Verify
oa --help
```

After this, `oa` is available from any directory.

To uninstall: `cargo uninstall office-automation`

## Test Sequence

Run these from the project root directory.

### 1. Basic checks (no Office needed)

```bash
oa --help
oa --version
oa config
oa clean -f
```

### 2. Inspect a template

```bash
oa info quick_test_files/rpm_2024_market_report_template.pptx
```

Expected: 30 slides, 86 OLE, 100 charts, 49 ntbl, 31 trns, 6 delt, 24 ccst, 3 templates found.

### 3. Full pipeline update

```bash
mkdir -p output

oa update quick_test_files/rpm_2024_market_report_template.pptx ^
  -e quick_test_files/rpm_tracking_Argentina_(05_07).xlsx ^
  -o output/test_argentina.pptx
```

Expected: ~10s, 86 links, 80 tables, 6 deltas, 24 coloring, 100 charts.

### 4. Verify the output

```bash
oa info output/test_argentina.pptx
```

Should match the template counts (same shapes, updated data).

### 5. Validate against Excel

```bash
oa check output/test_argentina.pptx -e quick_test_files/rpm_tracking_Argentina_(05_07).xlsx
```

Exit code 0 = all cells match. Exit code 1 = mismatches found (will list them).

### 6. Compare two files

```bash
oa diff quick_test_files/rpm_2024_market_report_template.pptx output/test_argentina.pptx
```

### 7. Test with different data

```bash
oa update quick_test_files/rpm_2024_market_report_template.pptx ^
  -e quick_test_files/rpm_tracking_Mexico_(05_07).xlsx ^
  -o output/test_mexico.pptx

oa update quick_test_files/rpm_2024_market_report_template.pptx ^
  -e quick_test_files/rpm_tracking_United_States_(05_07).xlsx ^
  -o output/test_us.pptx
```

### 8. Selective steps

```bash
# Only update tables (skip links, deltas, coloring, charts)
oa update output/test_argentina.pptx ^
  -e quick_test_files/rpm_tracking_Argentina_(05_07).xlsx ^
  --steps tables

# Skip charts (everything else runs)
oa update quick_test_files/rpm_2024_market_report_template.pptx ^
  -e quick_test_files/rpm_tracking_Argentina_(05_07).xlsx ^
  -o output/test_skip_charts.pptx ^
  --skip charts
```

### 9. Dry run

```bash
oa update quick_test_files/rpm_2024_market_report_template.pptx ^
  -e quick_test_files/rpm_tracking_Argentina_(05_07).xlsx ^
  --dry-run
```

Should NOT modify the original file.

### 10. Test with the 71-slide template

```bash
oa info quick_test_files/test_template.pptx

oa update quick_test_files/test_template.pptx ^
  -e quick_test_files/rpm_tracking_United_States_(05_07).xlsx ^
  -o output/test_71slides.pptx
```

Expected: 71 slides, 150 OLE, 257 charts.

## Troubleshooting

**"COM error: Exception occurred"**
- Run `oa clean -f` to kill zombie Office processes, then retry

**"File not found"**
- Use absolute paths or run from the project root

**"RPC server is unavailable"**
- Excel or PowerPoint crashed mid-run. Run `oa clean -f` and retry

**Slow performance (>20s)**
- Ensure no other Office instances are open
- Check `oa clean` for zombies first

## Running Automated Tests

```bash
# Unit tests (no Office needed)
cargo test

# Integration tests (needs Office)
OA_INTEGRATION=1 cargo test --test com_smoke -- --ignored --test-threads=1
```
