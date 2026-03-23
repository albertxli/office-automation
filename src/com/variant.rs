//! Variant newtype over `windows::Win32::System::Variant::VARIANT`.
//!
//! Provides ergonomic conversions for COM automation values.

use windows::core::BSTR;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::*;

use crate::error::{OaError, OaResult};

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
}
