//! RAII guards for COM lifecycle management.
//!
//! Ensures COM is properly initialized and deinitialized, and Office apps
//! are quit in the correct order to prevent zombie processes and 60-second hangs.

use windows::Win32::System::Com::CoUninitialize;

/// RAII guard that calls `CoUninitialize` on drop.
///
/// Create one per thread that uses COM. Must be the LAST thing dropped
/// (i.e., created before any COM objects).
pub struct ComGuard {
    _private: (), // prevent construction outside this module
}

impl ComGuard {
    /// Create a new ComGuard. Caller must have already called `CoInitializeEx`.
    pub(crate) fn new() -> Self {
        Self { _private: () }
    }
}

impl Drop for ComGuard {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}

/// Helper to ensure correct drop ordering for an Office COM session.
///
/// The critical ordering (from GOTCHA #21) is:
/// 1. Drop all shape/inventory references (IDispatch pointers)
/// 2. Drop the Presentation object
/// 3. Call Excel Application.Quit()
/// 4. Call PowerPoint Application.Quit()
/// 5. Drop the ComGuard (calls CoUninitialize)
///
/// Violating this order causes a 60-second hang waiting for RPC disconnection.
///
/// Usage:
/// ```ignore
/// let _com = ComGuard::new();
/// // ... create apps, do work ...
/// // Explicit cleanup:
/// drop(inventory);
/// drop(presentation);
/// excel_app.call0("Quit")?;
/// ppt_app.call0("Quit")?;
/// // _com drops here, calling CoUninitialize
/// ```
///
/// This struct documents the pattern — actual cleanup is done manually since
/// the types involved (Dispatch, Variant) don't form a single owning hierarchy.
#[allow(dead_code)]
pub struct SessionCleanup;

#[allow(dead_code)]
impl SessionCleanup {
    /// Document: correct cleanup order for reference.
    pub const CLEANUP_ORDER: &'static str =
        "1. drop(inventory) → 2. drop(presentation) → 3. excel.Quit() → 4. ppt.Quit() → 5. CoUninitialize";
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cleanup_order_documented() {
        assert!(SessionCleanup::CLEANUP_ORDER.contains("inventory"));
        assert!(SessionCleanup::CLEANUP_ORDER.contains("presentation"));
        assert!(SessionCleanup::CLEANUP_ORDER.contains("Quit"));
        assert!(SessionCleanup::CLEANUP_ORDER.contains("CoUninitialize"));
    }
}
