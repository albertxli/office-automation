//! Integration test: ZIP relink.
//!
//! Copies a PPTX, relinks it to a different Excel file, then verifies
//! the links were changed using `oa info`.

use std::path::PathBuf;

use office_automation::zip_ops::detector::detect_linked_excel;
use office_automation::zip_ops::relinker::relink_pptx_zip;

fn test_data_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("quick_test_files")
}

#[test]
fn test_detect_linked_excel() {
    let pptx = test_data_dir().join("test_template.pptx");
    if !pptx.exists() {
        eprintln!("Skipping: test_template.pptx not found");
        return;
    }

    let result = detect_linked_excel(&pptx);
    assert!(result.is_some(), "Expected to find linked Excel");
    let excel_path = result.unwrap();
    let path_str = excel_path.to_string_lossy().to_lowercase();
    assert!(path_str.ends_with(".xlsx"), "Expected .xlsx path, got: {path_str}");
    println!("Detected Excel: {}", excel_path.display());
}

#[test]
fn test_relink_pptx_zip() {
    let pptx_src = test_data_dir().join("test_template.pptx");
    // Use Mexico Excel as target — different from whatever the template currently points to
    let excel_target = test_data_dir().join("rpm_tracking_Mexico_(05_07).xlsx");

    if !pptx_src.exists() || !excel_target.exists() {
        eprintln!("Skipping: test files not found");
        return;
    }

    // Copy PPTX to temp location
    let tmp_dir = tempfile::tempdir().unwrap();
    let tmp_pptx = tmp_dir.path().join("test_relinked.pptx");
    std::fs::copy(&pptx_src, &tmp_pptx).unwrap();

    // Detect original Excel link
    let original = detect_linked_excel(&tmp_pptx);
    assert!(original.is_some(), "No links found before relink");
    let original_path = original.unwrap();
    println!("Original link: {}", original_path.display());

    // First relink to Mexico (guaranteed different from whatever it is now)
    let count = relink_pptx_zip(&tmp_pptx, &excel_target).unwrap();
    // If original was already Mexico, count=0 is OK. Do a second relink to US to guarantee change.
    if count == 0 {
        let us_excel = test_data_dir().join("rpm_tracking_United_States_(05_07).xlsx");
        if us_excel.exists() {
            let count2 = relink_pptx_zip(&tmp_pptx, &us_excel).unwrap();
            assert!(count2 > 0, "Expected at least 1 link rewritten on second attempt");
            println!("Relinked {count2} links (second pass to US)");
        }
    } else {
        println!("Relinked {count} links to Mexico");
    }

    // Verify the link changed from original
    let new_link = detect_linked_excel(&tmp_pptx);
    assert!(new_link.is_some(), "No links found after relink");
    let new_path = new_link.unwrap();
    println!("New link: {}", new_path.display());
    assert_ne!(
        original_path.to_string_lossy().to_lowercase(),
        new_path.to_string_lossy().to_lowercase(),
        "Link should have changed after relink"
    );
}
