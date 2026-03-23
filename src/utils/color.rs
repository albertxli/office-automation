//! Color conversion utilities for COM automation.
//!
//! COM/VBA uses BGR-encoded Long values (B*65536 + G*256 + R).
//! These functions convert between `#RRGGBB` hex strings and BGR Longs.

/// Convert `#RRGGBB` hex string to a COM-compatible BGR Long value.
///
/// VBA's `RGB()` returns `R + (G * 256) + (B * 65536)` (BGR order as Long).
/// Returns 0 (black) for invalid input.
pub fn hex_to_bgr(hex_color: &str) -> i32 {
    let hex = hex_color.trim_start_matches('#');
    if hex.len() != 6 {
        return 0; // black fallback
    }

    let r = i32::from_str_radix(&hex[0..2], 16).unwrap_or(0);
    let g = i32::from_str_radix(&hex[2..4], 16).unwrap_or(0);
    let b = i32::from_str_radix(&hex[4..6], 16).unwrap_or(0);

    r + (g * 256) + (b * 65536)
}

/// Convert a BGR Long value back to `#RRGGBB` hex string.
pub fn bgr_to_hex(bgr: i32) -> String {
    let r = bgr & 0xFF;
    let g = (bgr >> 8) & 0xFF;
    let b = (bgr >> 16) & 0xFF;
    format!("#{:02X}{:02X}{:02X}", r, g, b)
}

/// Choose a dark or light font color based on background brightness.
///
/// Uses the weighted luminance formula: `0.299*R + 0.587*G + 0.114*B`.
/// Returns `light_bgr` if brightness < 128, otherwise `dark_bgr`.
///
/// `bg_color` is a BGR Long value (same as COM uses).
pub fn contrast_font_color(bg_color: i32, dark_bgr: i32, light_bgr: i32) -> i32 {
    let r = (bg_color & 0xFF) as f64;
    let g = ((bg_color >> 8) & 0xFF) as f64;
    let b = ((bg_color >> 16) & 0xFF) as f64;

    let brightness = 0.299 * r + 0.587 * g + 0.114 * b;

    if brightness < 128.0 {
        light_bgr
    } else {
        dark_bgr
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hex_to_bgr_red() {
        // #FF0000 (red) → R=255, G=0, B=0 → 255
        assert_eq!(hex_to_bgr("#FF0000"), 255);
    }

    #[test]
    fn test_hex_to_bgr_green() {
        // #00FF00 (green) → R=0, G=255, B=0 → 65280
        assert_eq!(hex_to_bgr("#00FF00"), 255 * 256);
    }

    #[test]
    fn test_hex_to_bgr_blue() {
        // #0000FF (blue) → R=0, G=0, B=255 → 16711680
        assert_eq!(hex_to_bgr("#0000FF"), 255 * 65536);
    }

    #[test]
    fn test_hex_to_bgr_white() {
        assert_eq!(hex_to_bgr("#FFFFFF"), 255 + 255 * 256 + 255 * 65536);
    }

    #[test]
    fn test_hex_to_bgr_black() {
        assert_eq!(hex_to_bgr("#000000"), 0);
    }

    #[test]
    fn test_hex_to_bgr_no_hash() {
        assert_eq!(hex_to_bgr("FF0000"), 255);
    }

    #[test]
    fn test_hex_to_bgr_invalid_short() {
        assert_eq!(hex_to_bgr("#FFF"), 0);
    }

    #[test]
    fn test_hex_to_bgr_invalid_long() {
        assert_eq!(hex_to_bgr("#FFFFFFF"), 0);
    }

    #[test]
    fn test_hex_to_bgr_lowercase() {
        assert_eq!(hex_to_bgr("#ff0000"), 255);
    }

    #[test]
    fn test_bgr_to_hex_red() {
        assert_eq!(bgr_to_hex(255), "#FF0000");
    }

    #[test]
    fn test_bgr_to_hex_green() {
        assert_eq!(bgr_to_hex(255 * 256), "#00FF00");
    }

    #[test]
    fn test_bgr_to_hex_blue() {
        assert_eq!(bgr_to_hex(255 * 65536), "#0000FF");
    }

    #[test]
    fn test_hex_bgr_roundtrip() {
        let colors = ["#F8696B", "#FFEB84", "#63BE7B", "#33CC33", "#ED0590", "#595959"];
        for hex in colors {
            assert_eq!(bgr_to_hex(hex_to_bgr(hex)), hex, "Roundtrip failed for {hex}");
        }
    }

    #[test]
    fn test_contrast_dark_bg_returns_light() {
        let dark_bg = hex_to_bgr("#000000");
        let dark_font = hex_to_bgr("#000000");
        let light_font = hex_to_bgr("#FFFFFF");
        assert_eq!(contrast_font_color(dark_bg, dark_font, light_font), light_font);
    }

    #[test]
    fn test_contrast_light_bg_returns_dark() {
        let light_bg = hex_to_bgr("#FFFFFF");
        let dark_font = hex_to_bgr("#000000");
        let light_font = hex_to_bgr("#FFFFFF");
        assert_eq!(contrast_font_color(light_bg, dark_font, light_font), dark_font);
    }

    #[test]
    fn test_contrast_mid_gray() {
        // Gray #808080: R=128, G=128, B=128
        // Brightness = 0.299*128 + 0.587*128 + 0.114*128
        // Due to IEEE 754 float precision (GOTCHA #7), this computes to
        // 127.99999999999999 (not exactly 128.0), so it falls below threshold.
        // This matches the Python behavior.
        let mid_gray = hex_to_bgr("#808080");
        let dark_font = 0;
        let light_font = 1;
        assert_eq!(contrast_font_color(mid_gray, dark_font, light_font), light_font);
    }

    #[test]
    fn test_contrast_just_below_threshold() {
        // #7F7F7F: R=127, G=127, B=127
        // Brightness = 0.299*127 + 0.587*127 + 0.114*127 = 127.0
        // 127.0 < 128, so should return light font
        let dark_gray = hex_to_bgr("#7F7F7F");
        let dark_font = 0;
        let light_font = 1;
        assert_eq!(contrast_font_color(dark_gray, dark_font, light_font), light_font);
    }

    #[test]
    fn test_contrast_config_defaults() {
        // Test with actual config default colors
        let bg = hex_to_bgr("#F8696B"); // heatmap red
        let dark = hex_to_bgr("#000000");
        let light = hex_to_bgr("#FFFFFF");
        // This is a fairly bright red, should get dark font
        let result = contrast_font_color(bg, dark, light);
        assert_eq!(result, dark);
    }
}
