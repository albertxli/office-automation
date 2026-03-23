//! COM session management for Office automation.
//!
//! Handles COM initialization (STA), Office application creation via
//! `CoCreateInstance`, and the background security dialog dismisser thread.

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use std::thread;
use std::time::Duration;

use windows::core::PCWSTR;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::System::Com::*;
use windows::Win32::UI::WindowsAndMessaging::*;

use crate::com::cleanup::ComGuard;
use crate::com::dispatch::Dispatch;
use crate::error::{OaError, OaResult};

/// Initialize COM on the current thread with Single-Threaded Apartment model.
///
/// Office COM requires STA — using MTA will cause mysterious failures.
/// Returns a `ComGuard` that calls `CoUninitialize` on drop.
pub fn init_com_sta() -> OaResult<ComGuard> {
    unsafe {
        let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        if hr.is_err() {
            return Err(OaError::Com(hr.into()));
        }
    }
    Ok(ComGuard::new())
}

/// Create a COM instance from a ProgID string (e.g., "Excel.Application").
///
/// Uses `CLSCTX_LOCAL_SERVER` for process isolation, equivalent to Python's
/// `win32com.client.DispatchEx` which creates a new, dedicated process.
pub fn create_instance(prog_id: &str) -> OaResult<Dispatch> {
    unsafe {
        // Convert ProgID to CLSID
        let wide_prog_id: Vec<u16> = prog_id.encode_utf16().chain(std::iter::once(0)).collect();
        let clsid = CLSIDFromProgID(PCWSTR(wide_prog_id.as_ptr()))?;

        // Create the COM object in a local server (separate process)
        let dispatch: IDispatch =
            CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)?;

        Ok(Dispatch::new(dispatch))
    }
}

/// Spawn a background thread that auto-dismisses the PowerPoint security dialog.
///
/// PowerPoint shows a "Microsoft PowerPoint Security Notice" when opening files
/// with OLE links. This dialog blocks COM automation until dismissed.
///
/// The thread polls for the dialog window every 500ms and sends WM_CLOSE to
/// dismiss it. Returns a stop flag and join handle.
///
/// # Returns
/// `(stop_flag, join_handle)` — set `stop_flag` to `true` to stop the thread.
pub fn spawn_dialog_dismisser() -> (Arc<AtomicBool>, thread::JoinHandle<()>) {
    let stop = Arc::new(AtomicBool::new(false));
    let stop_clone = stop.clone();

    let handle = thread::spawn(move || {
        // This thread needs its own COM initialization for FindWindowW
        // (though FindWindowW doesn't strictly require COM, being safe)
        // Poll more aggressively — the dialog can appear and block within 100ms
        while !stop_clone.load(Ordering::Relaxed) {
            dismiss_security_dialog();
            thread::sleep(Duration::from_millis(100));
        }
    });

    (stop, handle)
}

/// Try to find and dismiss the PowerPoint security dialog.
fn dismiss_security_dialog() {
    // The dialog title varies by Office version/language.
    // Common: "Microsoft PowerPoint Security Notice"
    let titles = [
        "Microsoft PowerPoint Security Notice",
        "Microsoft PowerPoint",
    ];

    for title in &titles {
        let wide_title: Vec<u16> = title.encode_utf16().chain(std::iter::once(0)).collect();

        unsafe {
            if let Ok(hwnd) = FindWindowW(None, PCWSTR(wide_title.as_ptr())) {
                if hwnd != HWND::default() {
                    // Found the dialog — send WM_CLOSE to dismiss it
                    let _ = PostMessageW(Some(hwnd), WM_CLOSE, WPARAM(0), LPARAM(0));
                }
            }
        }
    }
}

/// Stop the dialog dismisser thread and wait for it to finish.
pub fn stop_dialog_dismisser(
    stop: Arc<AtomicBool>,
    handle: thread::JoinHandle<()>,
) {
    stop.store(true, Ordering::Relaxed);
    let _ = handle.join();
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_spawn_and_stop_dismisser() {
        // Verify the thread starts and stops cleanly
        let (stop, handle) = spawn_dialog_dismisser();
        // Let it run briefly
        thread::sleep(Duration::from_millis(100));
        stop_dialog_dismisser(stop, handle);
        // If we get here without panic, the thread lifecycle works
    }

    #[test]
    fn test_dismiss_security_dialog_no_crash() {
        // Should not crash even if no dialog exists
        dismiss_security_dialog();
    }
}
