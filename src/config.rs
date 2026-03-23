use serde::Deserialize;

use crate::error::{OaError, OaResult};

/// Heatmap color configuration for 3-color scale tables (htmp_ shapes).
#[derive(Debug, Clone, Deserialize)]
pub struct HeatmapConfig {
    pub color_minimum: String,
    pub color_midpoint: String,
    pub color_maximum: String,
    pub dark_font: String,
    pub light_font: String,
}

impl Default for HeatmapConfig {
    fn default() -> Self {
        Self {
            color_minimum: "#F8696B".into(),
            color_midpoint: "#FFEB84".into(),
            color_maximum: "#63BE7B".into(),
            dark_font: "#000000".into(),
            light_font: "#FFFFFF".into(),
        }
    }
}

/// Color coding configuration for _ccst shapes (sign-based coloring).
#[derive(Debug, Clone, Deserialize)]
pub struct CcstConfig {
    pub positive_color: String,
    pub negative_color: String,
    pub neutral_color: String,
    pub positive_prefix: String,
    pub symbol_removal: String,
}

impl Default for CcstConfig {
    fn default() -> Self {
        Self {
            positive_color: "#33CC33".into(),
            negative_color: "#ED0590".into(),
            neutral_color: "#595959".into(),
            positive_prefix: "+".into(),
            symbol_removal: "%".into(),
        }
    }
}

/// Delta indicator configuration (template shape names and source slide).
#[derive(Debug, Clone, Deserialize)]
pub struct DeltaConfig {
    pub template_positive: String,
    pub template_negative: String,
    pub template_none: String,
    pub template_slide: i32,
}

impl Default for DeltaConfig {
    fn default() -> Self {
        Self {
            template_positive: "tmpl_delta_pos".into(),
            template_negative: "tmpl_delta_neg".into(),
            template_none: "tmpl_delta_none".into(),
            template_slide: 1,
        }
    }
}

/// OLE link behavior configuration.
#[derive(Debug, Clone, Deserialize)]
pub struct LinksConfig {
    pub set_manual: bool,
}

impl Default for LinksConfig {
    fn default() -> Self {
        Self { set_manual: true }
    }
}

/// Top-level configuration with all sections.
#[derive(Debug, Clone, Deserialize, Default)]
pub struct Config {
    pub heatmap: HeatmapConfig,
    pub ccst: CcstConfig,
    pub delta: DeltaConfig,
    pub links: LinksConfig,
}

impl Config {
    /// Apply `--set KEY=VALUE` overrides using dot notation.
    ///
    /// Keys are like `heatmap.color_minimum`, `ccst.positive_color`, etc.
    pub fn apply_overrides(&mut self, overrides: &[String]) -> OaResult<()> {
        for item in overrides {
            let (key, value) = item
                .split_once('=')
                .ok_or_else(|| OaError::Config(format!("Invalid --set format: {item:?} (expected KEY=VALUE)")))?;

            let key = key.trim();
            let value = value.trim();

            match key {
                // Heatmap
                "heatmap.color_minimum" => self.heatmap.color_minimum = value.into(),
                "heatmap.color_midpoint" => self.heatmap.color_midpoint = value.into(),
                "heatmap.color_maximum" => self.heatmap.color_maximum = value.into(),
                "heatmap.dark_font" => self.heatmap.dark_font = value.into(),
                "heatmap.light_font" => self.heatmap.light_font = value.into(),
                // CCST
                "ccst.positive_color" => self.ccst.positive_color = value.into(),
                "ccst.negative_color" => self.ccst.negative_color = value.into(),
                "ccst.neutral_color" => self.ccst.neutral_color = value.into(),
                "ccst.positive_prefix" => self.ccst.positive_prefix = value.into(),
                "ccst.symbol_removal" => self.ccst.symbol_removal = value.into(),
                // Delta
                "delta.template_positive" => self.delta.template_positive = value.into(),
                "delta.template_negative" => self.delta.template_negative = value.into(),
                "delta.template_none" => self.delta.template_none = value.into(),
                "delta.template_slide" => {
                    self.delta.template_slide = value.parse::<i32>().map_err(|_| {
                        OaError::Config(format!("Invalid integer for delta.template_slide: {value:?}"))
                    })?;
                }
                // Links
                "links.set_manual" => {
                    self.links.set_manual = coerce_bool(value).ok_or_else(|| {
                        OaError::Config(format!("Invalid boolean for links.set_manual: {value:?}"))
                    })?;
                }
                _ => {
                    return Err(OaError::Config(format!("Unknown config key: {key:?}")));
                }
            }
        }
        Ok(())
    }

    /// All valid config keys and their current values, for `oa config`.
    pub fn all_keys(&self) -> Vec<(&'static str, String)> {
        vec![
            ("heatmap.color_minimum", self.heatmap.color_minimum.clone()),
            ("heatmap.color_midpoint", self.heatmap.color_midpoint.clone()),
            ("heatmap.color_maximum", self.heatmap.color_maximum.clone()),
            ("heatmap.dark_font", self.heatmap.dark_font.clone()),
            ("heatmap.light_font", self.heatmap.light_font.clone()),
            ("ccst.positive_color", self.ccst.positive_color.clone()),
            ("ccst.negative_color", self.ccst.negative_color.clone()),
            ("ccst.neutral_color", self.ccst.neutral_color.clone()),
            ("ccst.positive_prefix", self.ccst.positive_prefix.clone()),
            ("ccst.symbol_removal", self.ccst.symbol_removal.clone()),
            ("delta.template_positive", self.delta.template_positive.clone()),
            ("delta.template_negative", self.delta.template_negative.clone()),
            ("delta.template_none", self.delta.template_none.clone()),
            ("delta.template_slide", self.delta.template_slide.to_string()),
            ("links.set_manual", self.links.set_manual.to_string()),
        ]
    }
}

fn coerce_bool(s: &str) -> Option<bool> {
    match s.to_lowercase().as_str() {
        "true" | "1" | "yes" => Some(true),
        "false" | "0" | "no" => Some(false),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_config_has_all_15_keys() {
        let config = Config::default();
        assert_eq!(config.all_keys().len(), 15);
    }

    #[test]
    fn test_default_values() {
        let config = Config::default();
        assert_eq!(config.heatmap.color_minimum, "#F8696B");
        assert_eq!(config.ccst.positive_color, "#33CC33");
        assert_eq!(config.delta.template_positive, "tmpl_delta_pos");
        assert_eq!(config.delta.template_slide, 1);
        assert!(config.links.set_manual);
    }

    #[test]
    fn test_apply_overrides_string() {
        let mut config = Config::default();
        config
            .apply_overrides(&["ccst.positive_color=#FF0000".into()])
            .unwrap();
        assert_eq!(config.ccst.positive_color, "#FF0000");
    }

    #[test]
    fn test_apply_overrides_bool() {
        let mut config = Config::default();
        config
            .apply_overrides(&["links.set_manual=false".into()])
            .unwrap();
        assert!(!config.links.set_manual);
    }

    #[test]
    fn test_apply_overrides_int() {
        let mut config = Config::default();
        config
            .apply_overrides(&["delta.template_slide=3".into()])
            .unwrap();
        assert_eq!(config.delta.template_slide, 3);
    }

    #[test]
    fn test_apply_overrides_unknown_key() {
        let mut config = Config::default();
        let result = config.apply_overrides(&["unknown.key=value".into()]);
        assert!(result.is_err());
    }

    #[test]
    fn test_apply_overrides_bad_format() {
        let mut config = Config::default();
        let result = config.apply_overrides(&["no_equals_sign".into()]);
        assert!(result.is_err());
    }

    #[test]
    fn test_apply_overrides_bad_int() {
        let mut config = Config::default();
        let result = config.apply_overrides(&["delta.template_slide=abc".into()]);
        assert!(result.is_err());
    }

    #[test]
    fn test_apply_overrides_bad_bool() {
        let mut config = Config::default();
        let result = config.apply_overrides(&["links.set_manual=maybe".into()]);
        assert!(result.is_err());
    }

    #[test]
    fn test_apply_multiple_overrides() {
        let mut config = Config::default();
        config
            .apply_overrides(&[
                "heatmap.color_minimum=#000000".into(),
                "heatmap.color_maximum=#FFFFFF".into(),
                "ccst.positive_prefix=".into(),
            ])
            .unwrap();
        assert_eq!(config.heatmap.color_minimum, "#000000");
        assert_eq!(config.heatmap.color_maximum, "#FFFFFF");
        assert_eq!(config.ccst.positive_prefix, "");
    }

    #[test]
    fn test_coerce_bool_variants() {
        assert_eq!(coerce_bool("true"), Some(true));
        assert_eq!(coerce_bool("True"), Some(true));
        assert_eq!(coerce_bool("TRUE"), Some(true));
        assert_eq!(coerce_bool("1"), Some(true));
        assert_eq!(coerce_bool("yes"), Some(true));
        assert_eq!(coerce_bool("false"), Some(false));
        assert_eq!(coerce_bool("False"), Some(false));
        assert_eq!(coerce_bool("0"), Some(false));
        assert_eq!(coerce_bool("no"), Some(false));
        assert_eq!(coerce_bool("maybe"), None);
        assert_eq!(coerce_bool(""), None);
    }
}
