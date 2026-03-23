//! Variant newtype over `windows::Win32::System::Variant::VARIANT`.
//!
//! Provides ergonomic conversions for COM automation values.

use windows::core::BSTR;
use windows::Win32::System::Com::{IDispatch, SAFEARRAY};
use windows::Win32::System::Ole::{
    SafeArrayAccessData, SafeArrayGetDim, SafeArrayGetElement, SafeArrayGetLBound,
    SafeArrayGetUBound, SafeArrayUnaccessData,
};
use windows::Win32::System::Variant::*;

use crate::error::{OaError, OaResult};

// VARENUM constants for SAFEARRAY element type detection.
const VT_ARRAY: u16 = 0x2000;
const VT_R8: u16 = 5;
const VT_VARIANT: u16 = 12;

/// A value extracted from a SAFEARRAY element (Range.Value2 or Series.Values).
#[derive(Debug, Clone)]
pub enum CellValue {
    F64(f64),
    I32(i32),
    Str(String),
    Empty,
}

impl CellValue {
    /// Convert to f64, treating empty/string as 0.0 (matches Python's chart behavior).
    pub fn to_f64(&self) -> f64 {
        match self {
            CellValue::F64(v) => *v,
            CellValue::I32(v) => *v as f64,
            CellValue::Str(s) => s.parse::<f64>().unwrap_or(0.0),
            CellValue::Empty => 0.0,
        }
    }
}

/// A wrapper around COM VARIANT that provides ergonomic Rust conversions.
///
/// VARIANT itself doesn't implement Debug, so we implement it manually.
#[derive(Clone)]
pub struct Variant(pub VARIANT);

impl Variant {
    /// Create an empty variant (VT_EMPTY).
    pub fn empty() -> Self {
        Self(VARIANT::default())
    }

    /// Check if this variant is empty (VT_EMPTY).
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }

    /// Get the variant type tag.
    pub fn vt(&self) -> u16 {
        // SAFETY: reading the type discriminant from the union.
        // VARIANT -> VARIANT_0 (union) -> VARIANT_0_0 (ManuallyDrop) -> vt (VARENUM)
        unsafe { self.0.Anonymous.Anonymous.vt.0 }
    }

    /// Try to extract an i32 value.
    pub fn as_i32(&self) -> OaResult<i32> {
        i32::try_from(&self.0).map_err(OaError::Com)
    }

    /// Try to extract an f64 value.
    pub fn as_f64(&self) -> OaResult<f64> {
        f64::try_from(&self.0).map_err(OaError::Com)
    }

    /// Try to extract a string value (from BSTR).
    pub fn as_string(&self) -> OaResult<String> {
        let bstr = BSTR::try_from(&self.0).map_err(OaError::Com)?;
        Ok(bstr.to_string())
    }

    /// Try to extract a bool value.
    pub fn as_bool(&self) -> OaResult<bool> {
        bool::try_from(&self.0).map_err(OaError::Com)
    }

    /// Try to extract an IDispatch COM object.
    pub fn as_dispatch(&self) -> OaResult<IDispatch> {
        IDispatch::try_from(&self.0).map_err(OaError::Com)
    }

    /// Try to coerce to a numeric value (i32 or f64 → f64).
    pub fn as_numeric(&self) -> OaResult<f64> {
        // Try f64 first, then i32
        if let Ok(v) = self.as_f64() {
            return Ok(v);
        }
        if let Ok(v) = self.as_i32() {
            return Ok(v as f64);
        }
        Err(OaError::Other(format!("Cannot convert variant (vt={}) to numeric", self.vt())))
    }

    /// Get the inner VARIANT reference for passing to COM calls.
    pub fn inner(&self) -> &VARIANT {
        &self.0
    }

    /// Get a mutable reference to the inner VARIANT.
    pub fn inner_mut(&mut self) -> &mut VARIANT {
        &mut self.0
    }

    /// Consume and return the inner VARIANT.
    pub fn into_inner(self) -> VARIANT {
        self.0
    }

    /// Check if this variant contains a SAFEARRAY.
    pub fn is_array(&self) -> bool {
        self.vt() & VT_ARRAY != 0
    }

    /// Extract a 1D SAFEARRAY of f64 values (VT_ARRAY|VT_R8).
    ///
    /// Used for `Series.Values` which returns chart data points as doubles.
    /// Uses `SafeArrayAccessData` for zero-copy pointer access.
    pub fn as_f64_array(&self) -> OaResult<Vec<f64>> {
        unsafe {
            let psa = self.safearray_ptr()?;
            let dims = SafeArrayGetDim(psa);
            if dims != 1 {
                return Err(OaError::Other(format!("Expected 1D SAFEARRAY, got {dims}D")));
            }

            let lb = SafeArrayGetLBound(psa, 1).map_err(OaError::Com)?;
            let ub = SafeArrayGetUBound(psa, 1).map_err(OaError::Com)?;
            let count = (ub - lb + 1) as usize;
            if count == 0 {
                return Ok(vec![]);
            }

            let mut data_ptr: *mut std::ffi::c_void = std::ptr::null_mut();
            SafeArrayAccessData(psa, &mut data_ptr).map_err(OaError::Com)?;

            // Scope guard: always call UnaccessData even if we panic/error
            let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
                let f64_ptr = data_ptr as *const f64;
                let slice = std::slice::from_raw_parts(f64_ptr, count);
                slice.to_vec()
            }));

            SafeArrayUnaccessData(psa).map_err(OaError::Com)?;

            result.map_err(|_| OaError::Other("Panic while reading SAFEARRAY data".into()))
        }
    }

    /// Extract a flat Vec<f64> from any numeric VARIANT — scalar or array.
    ///
    /// Handles:
    /// - Scalar (VT_R8, VT_I4, VT_EMPTY) → vec![value]
    /// - 1D VT_ARRAY|VT_R8 → direct f64 array read
    /// - 1D/2D VT_ARRAY|VT_VARIANT → per-element extraction, flattened row-by-row
    pub fn as_flat_f64_vec(&self) -> OaResult<Vec<f64>> {
        if !self.is_array() {
            // Scalar: wrap in single-element vec
            if self.is_empty() {
                return Ok(vec![0.0]);
            }
            return Ok(vec![self.as_numeric()?]);
        }

        let elem_vt = self.vt() & 0x0FFF;

        if elem_vt == VT_R8 {
            // Fast path: 1D array of f64 (Series.Values typical case)
            return self.as_f64_array();
        }

        if elem_vt == VT_VARIANT {
            // Slow path: array of VARIANTs (Range.Value2 typical case)
            return self.as_variant_array_flat();
        }

        Err(OaError::Other(format!("Unsupported SAFEARRAY element type: VT={elem_vt}")))
    }

    /// Extract a SAFEARRAY of VARIANTs (1D or 2D), flattened to Vec<f64>.
    ///
    /// Range.Value2 returns 2D (rows × cols). Series.Values may return
    /// VT_ARRAY|VT_VARIANT in some edge cases. Both are handled.
    fn as_variant_array_flat(&self) -> OaResult<Vec<f64>> {
        unsafe {
            let psa = self.safearray_ptr()?;
            let dims = SafeArrayGetDim(psa);

            match dims {
                1 => self.read_variant_array_1d(psa),
                2 => self.read_variant_array_2d(psa),
                _ => Err(OaError::Other(format!("Unsupported {dims}D SAFEARRAY"))),
            }
        }
    }

    /// Read 1D SAFEARRAY of VARIANTs.
    unsafe fn read_variant_array_1d(&self, psa: *const SAFEARRAY) -> OaResult<Vec<f64>> {
        let lb = unsafe { SafeArrayGetLBound(psa, 1).map_err(OaError::Com)? };
        let ub = unsafe { SafeArrayGetUBound(psa, 1).map_err(OaError::Com)? };
        let mut values = Vec::with_capacity((ub - lb + 1) as usize);

        for i in lb..=ub {
            let val = unsafe { self.get_variant_element(psa, &[i])? };
            values.push(variant_to_f64(&val));
        }
        Ok(values)
    }

    /// Read 2D SAFEARRAY of VARIANTs, flattened row-by-row.
    ///
    /// Range.Value2 returns (rows, cols) where dim 1 = rows, dim 2 = cols.
    unsafe fn read_variant_array_2d(&self, psa: *const SAFEARRAY) -> OaResult<Vec<f64>> {
        let row_lb = unsafe { SafeArrayGetLBound(psa, 1).map_err(OaError::Com)? };
        let row_ub = unsafe { SafeArrayGetUBound(psa, 1).map_err(OaError::Com)? };
        let col_lb = unsafe { SafeArrayGetLBound(psa, 2).map_err(OaError::Com)? };
        let col_ub = unsafe { SafeArrayGetUBound(psa, 2).map_err(OaError::Com)? };

        let rows = (row_ub - row_lb + 1) as usize;
        let cols = (col_ub - col_lb + 1) as usize;
        let mut values = Vec::with_capacity(rows * cols);

        for r in row_lb..=row_ub {
            for c in col_lb..=col_ub {
                let val = unsafe { self.get_variant_element(psa, &[r, c])? };
                values.push(variant_to_f64(&val));
            }
        }
        Ok(values)
    }

    /// Get a single VARIANT element from a SAFEARRAY by indices.
    ///
    /// SafeArrayGetElement copies the element — caller owns the result.
    unsafe fn get_variant_element(&self, psa: *const SAFEARRAY, indices: &[i32]) -> OaResult<VARIANT> {
        let mut element = VARIANT::default();
        unsafe {
            SafeArrayGetElement(
                psa,
                indices.as_ptr(),
                &mut element as *mut VARIANT as *mut std::ffi::c_void,
            )
            .map_err(OaError::Com)?;
        }
        Ok(element)
    }

    /// Get the SAFEARRAY pointer from this VARIANT.
    ///
    /// The VARIANT owns the SAFEARRAY — do NOT call SafeArrayDestroy on it.
    unsafe fn safearray_ptr(&self) -> OaResult<*const SAFEARRAY> {
        let psa = unsafe { self.0.Anonymous.Anonymous.Anonymous.parray };
        if psa.is_null() {
            return Err(OaError::Other("SAFEARRAY pointer is null".into()));
        }
        Ok(psa as *const SAFEARRAY)
    }
}

// --- From implementations ---

impl From<i32> for Variant {
    fn from(v: i32) -> Self {
        Self(VARIANT::from(v))
    }
}

impl From<f64> for Variant {
    fn from(v: f64) -> Self {
        Self(VARIANT::from(v))
    }
}

impl From<bool> for Variant {
    fn from(v: bool) -> Self {
        Self(VARIANT::from(v))
    }
}

impl From<&str> for Variant {
    fn from(v: &str) -> Self {
        Self(VARIANT::from(BSTR::from(v)))
    }
}

impl From<String> for Variant {
    fn from(v: String) -> Self {
        Self(VARIANT::from(BSTR::from(v.as_str())))
    }
}

impl From<BSTR> for Variant {
    fn from(v: BSTR) -> Self {
        Self(VARIANT::from(v))
    }
}

impl From<IDispatch> for Variant {
    fn from(v: IDispatch) -> Self {
        Self(VARIANT::from(v))
    }
}

impl From<VARIANT> for Variant {
    fn from(v: VARIANT) -> Self {
        Self(v)
    }
}

/// Convert a raw VARIANT element to f64, treating empty/error as 0.0.
///
/// Used when unpacking SAFEARRAY elements from Range.Value2.
fn variant_to_f64(v: &VARIANT) -> f64 {
    // Try f64 first (most common for numeric data)
    if let Ok(val) = f64::try_from(v) {
        return val;
    }
    // Try i32 (integer cells)
    if let Ok(val) = i32::try_from(v) {
        return val as f64;
    }
    // Empty/null/error/string → 0.0 (matches Python: empty cells plot as zero)
    0.0
}

impl Default for Variant {
    fn default() -> Self {
        Self::empty()
    }
}

impl std::fmt::Debug for Variant {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Variant")
            .field("vt", &self.vt())
            .finish()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_empty_variant() {
        let v = Variant::empty();
        assert!(v.is_empty());
    }

    #[test]
    fn test_i32_round_trip() {
        let v = Variant::from(42i32);
        assert_eq!(v.as_i32().unwrap(), 42);
    }

    #[test]
    fn test_f64_round_trip() {
        let v = Variant::from(3.14f64);
        assert!((v.as_f64().unwrap() - 3.14).abs() < f64::EPSILON);
    }

    #[test]
    fn test_bool_round_trip() {
        let v_true = Variant::from(true);
        let v_false = Variant::from(false);
        assert!(v_true.as_bool().unwrap());
        assert!(!v_false.as_bool().unwrap());
    }

    #[test]
    fn test_string_round_trip() {
        let v = Variant::from("hello world");
        assert_eq!(v.as_string().unwrap(), "hello world");
    }

    #[test]
    fn test_string_from_owned() {
        let v = Variant::from("test string".to_string());
        assert_eq!(v.as_string().unwrap(), "test string");
    }

    #[test]
    fn test_numeric_from_i32() {
        let v = Variant::from(10i32);
        assert!((v.as_numeric().unwrap() - 10.0).abs() < f64::EPSILON);
    }

    #[test]
    fn test_numeric_from_f64() {
        let v = Variant::from(2.5f64);
        assert!((v.as_numeric().unwrap() - 2.5).abs() < f64::EPSILON);
    }

    #[test]
    fn test_empty_coerces_to_zero() {
        // COM's VariantToDouble coerces VT_EMPTY to 0.0 — this is correct behavior.
        let v = Variant::empty();
        assert!((v.as_numeric().unwrap() - 0.0).abs() < f64::EPSILON);
    }

    #[test]
    fn test_default_is_empty() {
        let v = Variant::default();
        assert!(v.is_empty());
    }

    #[test]
    fn test_negative_i32() {
        let v = Variant::from(-1i32);
        assert_eq!(v.as_i32().unwrap(), -1);
    }

    #[test]
    fn test_zero_values() {
        let vi = Variant::from(0i32);
        let vf = Variant::from(0.0f64);
        assert_eq!(vi.as_i32().unwrap(), 0);
        assert!((vf.as_f64().unwrap()).abs() < f64::EPSILON);
    }

    #[test]
    fn test_empty_string() {
        let v = Variant::from("");
        assert_eq!(v.as_string().unwrap(), "");
    }

    #[test]
    fn test_unicode_string() {
        let v = Variant::from("日本語テスト");
        assert_eq!(v.as_string().unwrap(), "日本語テスト");
    }

    // --- CellValue tests ---

    #[test]
    fn test_cell_value_f64() {
        assert!((CellValue::F64(3.14).to_f64() - 3.14).abs() < f64::EPSILON);
    }

    #[test]
    fn test_cell_value_i32() {
        assert!((CellValue::I32(42).to_f64() - 42.0).abs() < f64::EPSILON);
    }

    #[test]
    fn test_cell_value_str_numeric() {
        assert!((CellValue::Str("2.5".into()).to_f64() - 2.5).abs() < f64::EPSILON);
    }

    #[test]
    fn test_cell_value_str_non_numeric() {
        assert!((CellValue::Str("N/A".into()).to_f64() - 0.0).abs() < f64::EPSILON);
    }

    #[test]
    fn test_cell_value_empty() {
        assert!((CellValue::Empty.to_f64() - 0.0).abs() < f64::EPSILON);
    }

    // --- is_array tests ---

    #[test]
    fn test_is_array_false_for_scalars() {
        assert!(!Variant::from(1i32).is_array());
        assert!(!Variant::from(1.0f64).is_array());
        assert!(!Variant::from("hello").is_array());
        assert!(!Variant::empty().is_array());
    }
}
