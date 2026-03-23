//! IDispatch wrapper with DISPID caching for COM late-binding automation.
//!
//! This is the core abstraction that replaces Python's `win32com.client.Dispatch`.
//! All Office COM operations go through this wrapper.

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use windows::core::{BSTR, GUID, PCWSTR};
use windows::Win32::System::Com::*;
use windows::Win32::System::Variant::VARIANT;

use crate::com::variant::Variant;
use crate::error::{OaError, OaResult};

/// IID_NULL — a zeroed GUID required by IDispatch::GetIDsOfNames and Invoke.
/// COM requires a non-null pointer to this; passing null causes RPC_X_NULL_REF_POINTER.
const IID_NULL: GUID = GUID::zeroed();

/// The named DISPID for property-put operations.
const DISPID_PROPERTYPUT: i32 = -3; // DISPID_PROPERTYPUT

/// Locale ID for COM calls (en-US).
const LCID_EN_US: u32 = 0x0409;

/// A wrapper around `IDispatch` that provides ergonomic property/method access
/// with per-instance DISPID caching via shared Rc<RefCell<HashMap>>.
///
/// # DISPID Caching
/// DISPIDs are per-COM-class (not global). However, when we clone a Dispatch, the
/// clone SHARES the same cache via Rc — so all references to the same COM object
/// benefit from each other's lookups. This is the key perf optimization vs Python.
pub struct Dispatch {
    inner: IDispatch,
    /// Shared DISPID cache — clones share this via Rc to avoid re-resolving names
    dispid_cache: Rc<RefCell<HashMap<String, i32>>>,
}

impl Dispatch {
    /// Create a new Dispatch wrapper around an IDispatch interface.
    pub fn new(inner: IDispatch) -> Self {
        Self {
            inner,
            dispid_cache: Rc::new(RefCell::new(HashMap::new())),
        }
    }

    /// Create a Dispatch that shares an existing DISPID cache.
    ///
    /// Use when wrapping COM objects of the same class (e.g., all Table Cell objects
    /// share the same DISPIDs for "Shape", "TextFrame", etc.). Avoids redundant
    /// GetIDsOfNames calls after the first object resolves each name.
    pub fn new_with_cache(inner: IDispatch, cache: Rc<RefCell<HashMap<String, i32>>>) -> Self {
        Self {
            inner,
            dispid_cache: cache,
        }
    }

    /// Get the DISPID cache for sharing with other Dispatch instances of the same COM class.
    pub fn cache(&self) -> Rc<RefCell<HashMap<String, i32>>> {
        self.dispid_cache.clone()
    }

    /// Get the underlying IDispatch reference (for passing to COM functions).
    pub fn as_raw(&self) -> &IDispatch {
        &self.inner
    }

    /// Consume and return the underlying IDispatch.
    pub fn into_raw(self) -> IDispatch {
        self.inner
    }

    /// Resolve a property/method name to its DISPID, using the shared per-instance cache.
    fn get_dispid(&mut self, name: &str) -> OaResult<i32> {
        // Check shared cache
        if let Some(&id) = self.dispid_cache.borrow().get(name) {
            return Ok(id);
        }

        // Cache miss — call GetIDsOfNames
        let wide_name = BSTR::from(name);
        let names = [PCWSTR(wide_name.as_ptr())];
        let mut dispid: i32 = 0;

        unsafe {
            self.inner.GetIDsOfNames(
                &IID_NULL,
                names.as_ptr(),
                1,
                LCID_EN_US,
                &mut dispid,
            )?;
        }

        self.dispid_cache.borrow_mut().insert(name.to_string(), dispid);
        Ok(dispid)
    }

    /// Get a property value by name.
    ///
    /// Equivalent to Python: `obj.PropertyName`
    pub fn get(&mut self, name: &str) -> OaResult<Variant> {
        let dispid = self.get_dispid(name)?;
        self.invoke_raw(dispid, DISPATCH_PROPERTYGET, &[])
    }

    /// Get a property value with arguments (indexed/parameterized property).
    ///
    /// Equivalent to Python: `obj.PropertyName(arg1, arg2)`
    pub fn get_with(&mut self, name: &str, args: &[Variant]) -> OaResult<Variant> {
        let dispid = self.get_dispid(name)?;
        self.invoke_raw(dispid, DISPATCH_PROPERTYGET, args)
    }

    /// Set a property value by name.
    ///
    /// Equivalent to Python: `obj.PropertyName = value`
    pub fn put(&mut self, name: &str, value: impl Into<Variant>) -> OaResult<()> {
        let dispid = self.get_dispid(name)?;
        let value = value.into();

        // For PROPERTYPUT, the value argument must be named with DISPID_PROPERTYPUT
        let mut args = [value.into_inner()];
        let mut named_args = [DISPID_PROPERTYPUT];

        let params = DISPPARAMS {
            rgvarg: args.as_mut_ptr(),
            rgdispidNamedArgs: named_args.as_mut_ptr(),
            cArgs: 1,
            cNamedArgs: 1,
        };

        let mut result = VARIANT::default();
        let mut excep = EXCEPINFO::default();

        unsafe {
            self.inner.Invoke(
                dispid,
                &IID_NULL,
                LCID_EN_US,
                DISPATCH_PROPERTYPUT,
                &params,
                Some(&mut result),
                Some(&mut excep),
                None,
            )?;
        }

        Ok(())
    }

    /// Call a method by name with arguments.
    ///
    /// Equivalent to Python: `obj.MethodName(arg1, arg2)`
    pub fn call(&mut self, name: &str, args: &[Variant]) -> OaResult<Variant> {
        let dispid = self.get_dispid(name)?;
        // Use DISPATCH_METHOD | DISPATCH_PROPERTYGET to handle methods that
        // can also be accessed as properties (common in Office COM)
        self.invoke_raw(
            dispid,
            DISPATCH_FLAGS(DISPATCH_METHOD.0 | DISPATCH_PROPERTYGET.0),
            args,
        )
    }

    /// Call a method with no arguments.
    ///
    /// Equivalent to Python: `obj.MethodName()`
    pub fn call0(&mut self, name: &str) -> OaResult<Variant> {
        self.call(name, &[])
    }

    /// Navigate a dotted path, returning the final dispatch object.
    ///
    /// `nav("Slides.Count")` is equivalent to `get("Slides")?.get("Count")?`
    /// but returns each intermediate result as a Dispatch.
    ///
    /// The last segment returns a Variant (not Dispatch), so this is mainly
    /// useful for intermediate navigation.
    pub fn nav(&mut self, path: &str) -> OaResult<Dispatch> {
        let segments: Vec<&str> = path.split('.').collect();
        if segments.is_empty() {
            return Err(OaError::Other("Empty navigation path".into()));
        }

        let mut current = self.get(&segments[0])?.as_dispatch()?;
        let mut current_dispatch = Dispatch::new(current);

        for &segment in &segments[1..] {
            current = current_dispatch.get(segment)?.as_dispatch()?;
            current_dispatch = Dispatch::new(current);
        }

        Ok(current_dispatch)
    }

    /// Low-level invoke with DISPID, flags, and arguments.
    fn invoke_raw(
        &self,
        dispid: i32,
        flags: DISPATCH_FLAGS,
        args: &[Variant],
    ) -> OaResult<Variant> {
        // COM expects arguments in reverse order
        let mut raw_args: Vec<VARIANT> = args.iter().rev().map(|a| a.0.clone()).collect();

        let params = DISPPARAMS {
            rgvarg: if raw_args.is_empty() {
                std::ptr::null_mut()
            } else {
                raw_args.as_mut_ptr()
            },
            rgdispidNamedArgs: std::ptr::null_mut(),
            cArgs: raw_args.len() as u32,
            cNamedArgs: 0,
        };

        let mut result = VARIANT::default();
        let mut excep = EXCEPINFO::default();
        let mut arg_err: u32 = 0;

        unsafe {
            self.inner
                .Invoke(
                    dispid,
                    &IID_NULL,
                    LCID_EN_US,
                    flags,
                    &params,
                    Some(&mut result),
                    Some(&mut excep),
                    Some(&mut arg_err),
                )
                .map_err(|e| {
                    // Try to extract a meaningful error from EXCEPINFO
                    if !excep.bstrDescription.is_empty() {
                        OaError::Com(windows::core::Error::new(
                            e.code(),
                            excep.bstrDescription.to_string(),
                        ))
                    } else {
                        OaError::Com(e)
                    }
                })?;
        }

        Ok(Variant::from(result))
    }
}

impl std::fmt::Debug for Dispatch {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Dispatch").finish()
    }
}

impl Clone for Dispatch {
    fn clone(&self) -> Self {
        Self {
            inner: self.inner.clone(),
            // Share the same cache — clones benefit from cached DISPIDs
            dispid_cache: self.dispid_cache.clone(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_dispid_propertyput_constant() {
        // DISPID_PROPERTYPUT is -3 per COM spec
        assert_eq!(DISPID_PROPERTYPUT, -3);
    }

    #[test]
    fn test_dispatch_flags_combine() {
        let combined = DISPATCH_FLAGS(DISPATCH_METHOD.0 | DISPATCH_PROPERTYGET.0);
        assert_eq!(combined.0, 0x1 | 0x2);
        assert_eq!(combined.0, 3);
    }

    #[test]
    fn test_lcid_en_us() {
        assert_eq!(LCID_EN_US, 0x0409);
    }
}
