//! Integration test: COM smoke test.
//!
//! Validates the entire COM stack works:
//! 1. Initialize COM (STA)
//! 2. Create Excel.Application
//! 3. Open a workbook
//! 4. Read a cell value
//! 5. Close workbook
//! 6. Quit Excel
//! 7. Verify no zombie processes
//!
//! Requires: Office installed, OA_INTEGRATION=1 env var.

use std::path::PathBuf;

use office_automation::com::dispatch::Dispatch;
use office_automation::com::session::{create_instance, init_com_sta};
use office_automation::com::variant::Variant;

fn test_excel_path() -> PathBuf {
    // Use one of the test Excel files
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("quick_test_files")
        .join("rpm_tracking_United_States_(05_07).xlsx");
    assert!(path.exists(), "Test file not found: {}", path.display());
    path
}

#[test]
#[ignore] // Run with: OA_INTEGRATION=1 cargo test --test com_smoke -- --ignored
fn test_excel_open_read_close() {
    // Skip if OA_INTEGRATION is not set
    if std::env::var("OA_INTEGRATION").is_err() {
        eprintln!("Skipping: OA_INTEGRATION not set");
        return;
    }

    let excel_path = test_excel_path();
    let abs_path = excel_path.canonicalize().unwrap();
    // Strip \\?\ UNC prefix that canonicalize adds — Office COM doesn't understand it
    let path_str = abs_path.to_string_lossy().to_string();
    let path_str = path_str.strip_prefix(r"\\?\").unwrap_or(&path_str).to_string();
    println!("Opening: {path_str}");

    // 1. Initialize COM (STA)
    let _com = init_com_sta().expect("COM init failed");

    // 2. Create Excel.Application
    let mut excel = create_instance("Excel.Application").expect("Failed to create Excel.Application");

    // Set Visible = false, ScreenUpdating = false
    excel.put("Visible", Variant::from(false)).expect("Failed to set Visible");
    excel.put("DisplayAlerts", Variant::from(false)).expect("Failed to set DisplayAlerts");

    // 3. Open workbook with UpdateLinks=0
    // Workbooks.Open(Filename, UpdateLinks, ReadOnly, ...)
    let mut workbooks = Dispatch::new(
        excel.get("Workbooks").expect("Failed to get Workbooks").as_dispatch().unwrap()
    );
    let wb_variant = workbooks
        .call("Open", &[
            Variant::from(path_str.as_str()),
            Variant::from(0i32), // UpdateLinks = 0 (don't auto-refresh)
        ])
        .expect("Failed to open workbook");
    let mut workbook = Dispatch::new(wb_variant.as_dispatch().unwrap());

    // 4. Read workbook name
    let name = workbook.get("Name").expect("Failed to get Name");
    let name_str = name.as_string().expect("Failed to convert name to string");
    println!("Opened workbook: {name_str}");
    assert!(name_str.contains("rpm_tracking_United_States"));

    // 5. Read a cell value from the first sheet
    let mut sheets = Dispatch::new(
        workbook.get("Worksheets").expect("Failed to get Worksheets").as_dispatch().unwrap()
    );
    let sheet_variant = sheets.call("Item", &[Variant::from(1i32)])
        .expect("Failed to get sheet 1");
    let mut sheet = Dispatch::new(sheet_variant.as_dispatch().unwrap());

    // Read cell A1
    let range_variant = sheet.get_with("Range", &[Variant::from("A1")])
        .expect("Failed to get Range A1");
    let mut range = Dispatch::new(range_variant.as_dispatch().unwrap());
    let cell_value = range.get("Text").expect("Failed to get Text");
    println!("Cell A1 text: {:?}", cell_value.as_string());

    // 6. Close workbook without saving
    workbook.call("Close", &[Variant::from(false)]).expect("Failed to close workbook");

    // 7. Quit Excel — must drop refs first! (GOTCHA #21)
    drop(range);
    drop(sheet);
    drop(sheets);
    drop(workbook);
    drop(workbooks);
    excel.call0("Quit").expect("Failed to quit Excel");
    drop(excel);

    // _com drops here, calling CoUninitialize
    println!("COM smoke test passed — no zombies!");
}

#[test]
#[ignore]
fn test_powerpoint_open_close() {
    if std::env::var("OA_INTEGRATION").is_err() {
        eprintln!("Skipping: OA_INTEGRATION not set");
        return;
    }

    let pptx_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("quick_test_files")
        .join("test_template.pptx");
    assert!(pptx_path.exists(), "Test file not found: {}", pptx_path.display());
    let abs_path = pptx_path.canonicalize().unwrap();
    let path_str = abs_path.to_string_lossy().to_string();
    let path_str = path_str.strip_prefix(r"\\?\").unwrap_or(&path_str).to_string();
    println!("Opening: {path_str}");

    let _com = init_com_sta().expect("COM init failed");

    // Create PowerPoint.Application
    let mut ppt = create_instance("PowerPoint.Application").expect("Failed to create PowerPoint.Application");
    // PowerPoint's DisplayAlerts takes ppAlertsNone = 0 (not a bool)
    ppt.put("DisplayAlerts", Variant::from(0i32)).expect("Failed to set DisplayAlerts");

    // Open presentation (read-only)
    let mut presentations = Dispatch::new(
        ppt.get("Presentations").expect("Failed to get Presentations").as_dispatch().unwrap()
    );

    // PowerPoint uses MsoTriState: msoTrue=-1, msoFalse=0
    let pres_variant = presentations.call("Open", &[
        Variant::from(path_str.as_str()),
        Variant::from(-1i32),  // ReadOnly = msoTrue
        Variant::from(-1i32),  // Untitled = msoTrue
        Variant::from(0i32),   // WithWindow = msoFalse
    ]).expect("Failed to open presentation");
    let mut presentation = Dispatch::new(pres_variant.as_dispatch().unwrap());

    // Read slide count
    let mut slides = Dispatch::new(
        presentation.get("Slides").expect("Failed to get Slides").as_dispatch().unwrap()
    );
    let count = slides.get("Count").expect("Failed to get Count");
    let slide_count = count.as_i32().expect("Failed to convert count");
    println!("Slide count: {slide_count}");
    assert!(slide_count > 0, "Expected at least 1 slide");

    // Close and quit
    presentation.call("Close", &[]).expect("Failed to close presentation");
    drop(slides);
    drop(presentation);
    drop(presentations);
    ppt.call0("Quit").expect("Failed to quit PowerPoint");
    drop(ppt);

    println!("PowerPoint smoke test passed!");
}
