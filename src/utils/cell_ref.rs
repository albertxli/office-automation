//! R1C1 ↔ A1 cell reference conversion.

/// Convert a single R1C1 cell reference like `R5C3` to A1 notation like `C5`.
fn convert_single_r1c1(cell: &str) -> Result<String, String> {
    let cell = cell.trim().to_uppercase();

    // Find 'C' after position 0 (first char is 'R')
    let c_pos = cell[1..]
        .find('C')
        .map(|p| p + 1)
        .ok_or_else(|| format!("Invalid R1C1 reference: {cell:?} (no 'C' found)"))?;

    let row_num: u32 = cell[1..c_pos]
        .parse()
        .map_err(|_| format!("Invalid row in R1C1 reference: {cell:?}"))?;

    let col_num: u32 = cell[c_pos + 1..]
        .parse()
        .map_err(|_| format!("Invalid column in R1C1 reference: {cell:?}"))?;

    let col_letter = col_number_to_letter(col_num);
    Ok(format!("{col_letter}{row_num}"))
}

/// Convert a column number (1-based) to Excel column letter(s).
///
/// 1 → A, 26 → Z, 27 → AA, 702 → ZZ, 703 → AAA
fn col_number_to_letter(mut n: u32) -> String {
    let mut result = String::new();
    while n > 0 {
        let remainder = ((n - 1) % 26) as u8;
        result.insert(0, (b'A' + remainder) as char);
        n = (n - 1) / 26;
    }
    result
}

/// Convert an R1C1 range string to A1 notation.
///
/// - `R1C1` → `A1`
/// - `R1C1:R5C5` → `A1:E5`
/// - `R1C27` → `AA1`
pub fn r1c1_to_a1(r1c1_range: &str) -> Result<String, String> {
    let parts: Vec<&str> = r1c1_range.split(':').collect();
    match parts.len() {
        1 => convert_single_r1c1(parts[0]),
        2 => {
            let start = convert_single_r1c1(parts[0])?;
            let end = convert_single_r1c1(parts[1])?;
            Ok(format!("{start}:{end}"))
        }
        _ => Err(format!("Invalid R1C1 range: {r1c1_range:?}")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_single_cell_a1() {
        assert_eq!(r1c1_to_a1("R1C1").unwrap(), "A1");
    }

    #[test]
    fn test_single_cell_c5() {
        assert_eq!(r1c1_to_a1("R5C3").unwrap(), "C5");
    }

    #[test]
    fn test_range() {
        assert_eq!(r1c1_to_a1("R1C1:R5C5").unwrap(), "A1:E5");
    }

    #[test]
    fn test_column_z() {
        assert_eq!(r1c1_to_a1("R1C26").unwrap(), "Z1");
    }

    #[test]
    fn test_column_aa() {
        assert_eq!(r1c1_to_a1("R1C27").unwrap(), "AA1");
    }

    #[test]
    fn test_column_az() {
        assert_eq!(r1c1_to_a1("R1C52").unwrap(), "AZ1");
    }

    #[test]
    fn test_column_ba() {
        assert_eq!(r1c1_to_a1("R1C53").unwrap(), "BA1");
    }

    #[test]
    fn test_column_zz() {
        assert_eq!(r1c1_to_a1("R1C702").unwrap(), "ZZ1");
    }

    #[test]
    fn test_column_aaa() {
        assert_eq!(r1c1_to_a1("R1C703").unwrap(), "AAA1");
    }

    #[test]
    fn test_large_range() {
        assert_eq!(r1c1_to_a1("R1C1:R100C30").unwrap(), "A1:AD100");
    }

    #[test]
    fn test_lowercase_input() {
        assert_eq!(r1c1_to_a1("r1c1").unwrap(), "A1");
    }

    #[test]
    fn test_whitespace_handling() {
        assert_eq!(r1c1_to_a1(" R1C1 ").unwrap(), "A1");
    }

    #[test]
    fn test_invalid_no_c() {
        assert!(r1c1_to_a1("R1").is_err());
    }

    #[test]
    fn test_invalid_no_row() {
        assert!(r1c1_to_a1("RC1").is_err());
    }

    #[test]
    fn test_col_number_to_letter() {
        assert_eq!(col_number_to_letter(1), "A");
        assert_eq!(col_number_to_letter(26), "Z");
        assert_eq!(col_number_to_letter(27), "AA");
        assert_eq!(col_number_to_letter(52), "AZ");
        assert_eq!(col_number_to_letter(53), "BA");
        assert_eq!(col_number_to_letter(702), "ZZ");
        assert_eq!(col_number_to_letter(703), "AAA");
    }
}
