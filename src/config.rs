use std::collections::BTreeMap;

use serde::Deserialize;

use crate::error::{OaError, OaResult};
use crate::shapes::matcher::is_exact_token_match;

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
    /// Global dead band (GOTCHA #45): `|value| < threshold` → "none". `0` = plain sign test.
    /// Decimal units, matching Excel `Value2` (a `2%` cell is `0.02`).
    pub threshold: f64,
    /// Per-category dead bands keyed by a whole token of the paired OLE object name,
    /// e.g. `globalnet` → covers `globalnet_pet` and `globalnet_dig`. Longest matching token wins.
    pub thresholds: BTreeMap<String, f64>,
}

impl Default for DeltaConfig {
    fn default() -> Self {
        Self {
            template_positive: "tmpl_delta_pos".into(),
            template_negative: "tmpl_delta_neg".into(),
            template_none: "tmpl_delta_none".into(),
            template_slide: 1,
            threshold: 0.0,
            thresholds: BTreeMap::new(),
        }
    }
}

impl DeltaConfig {
    /// Effective threshold for a delta paired with `ole_name`, and the token that supplied it
    /// (`None` = the global `delta.threshold`). Tokens match whole words only
    /// (`matcher::is_exact_token_match`); when several match, the longest wins, ties alphabetical.
    pub fn threshold_for(&self, ole_name: &str) -> (f64, Option<&str>) {
        let best = self.thresholds.iter()
            .filter(|(token, _)| is_exact_token_match(ole_name, token))
            .max_by(|(a, _), (b, _)| a.len().cmp(&b.len()).then_with(|| b.cmp(a)));
        match best {
            Some((token, value)) => (*value, Some(token.as_str())),
            None => (self.threshold, None),
        }
    }
}

/// Parse a threshold value: finite and `>= 0`.
fn parse_threshold(key: &str, value: &str) -> OaResult<f64> {
    match value.parse::<f64>() {
        Ok(v) if v.is_finite() && v >= 0.0 => Ok(v),
        _ => Err(OaError::Config(format!(
            "Invalid threshold for {key}: {value:?} (expected a number >= 0, e.g. 0.02)"
        ))),
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
                "delta.threshold" => {
                    self.delta.threshold = parse_threshold(key, value)?;
                }
                k if k.starts_with("delta.threshold.") => {
                    let token = k["delta.threshold.".len()..].trim();
                    if token.is_empty() {
                        return Err(OaError::Config(
                            "Empty token in delta.threshold.<token> (e.g. delta.threshold.globalnet=0.02)".into(),
                        ));
                    }
                    let v = parse_threshold(key, value)?;
                    self.delta.thresholds.insert(token.to_string(), v);
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
    /// Per-token `delta.threshold.<token>` entries are appended when present.
    pub fn all_keys(&self) -> Vec<(String, String)> {
        let mut keys: Vec<(String, String)> = vec![
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
            ("delta.threshold", self.delta.threshold.to_string()),
            ("links.set_manual", self.links.set_manual.to_string()),
        ]
        .into_iter()
        .map(|(k, v)| (k.to_string(), v))
        .collect();
        for (token, v) in &self.delta.thresholds {
            keys.push((format!("delta.threshold.{token}"), v.to_string()));
        }
        keys
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
    fn test_default_config_has_all_16_keys() {
        let config = Config::default();
        assert_eq!(config.all_keys().len(), 16);
        assert!(config.all_keys().iter().any(|(k, v)| k == "delta.threshold" && v == "0"));
    }

    // ── GOTCHA #45: delta thresholds ───────────────────────

    #[test]
    fn test_delta_threshold_global_override() {
        let mut config = Config::default();
        config.apply_overrides(&["delta.threshold=0.02".into()]).unwrap();
        assert_eq!(config.delta.threshold, 0.02);
        assert_eq!(config.delta.threshold_for("anything"), (0.02, None));
    }

    #[test]
    fn test_delta_threshold_per_token_override_and_listing() {
        let mut config = Config::default();
        config.apply_overrides(&[
            "delta.threshold.globalnet=0.02".into(),
            "delta.threshold.marketnet=0.05".into(),
            "delta.threshold.globalnet=0.03".into(), // later --set replaces earlier
        ]).unwrap();
        assert_eq!(config.delta.thresholds.len(), 2);
        assert_eq!(config.delta.thresholds["globalnet"], 0.03);
        let keys = config.all_keys();
        assert!(keys.iter().any(|(k, v)| k == "delta.threshold.globalnet" && v == "0.03"));
        assert!(keys.iter().any(|(k, v)| k == "delta.threshold.marketnet" && v == "0.05"));
    }

    #[test]
    fn test_delta_threshold_rejects_bad_values() {
        for bad in ["delta.threshold=-0.01", "delta.threshold=abc", "delta.threshold=NaN",
                    "delta.threshold.globalnet=-1", "delta.threshold.globalnet=x",
                    "delta.threshold.=0.02", "delta.thresholdx=0.02"] {
            let mut config = Config::default();
            assert!(config.apply_overrides(&[bad.into()]).is_err(), "{bad} should be rejected");
        }
    }

    #[test]
    fn test_threshold_for_whole_token_matching() {
        let mut config = Config::default();
        config.apply_overrides(&[
            "delta.threshold=0.01".into(),
            "delta.threshold.globalnet=0.02".into(),
            "delta.threshold.marketnet=0.05".into(),
        ]).unwrap();
        let d = &config.delta;
        // `_` is a token boundary: one token covers the whole category
        assert_eq!(d.threshold_for("globalnet_pet"), (0.02, Some("globalnet")));
        assert_eq!(d.threshold_for("globalnet_dig"), (0.02, Some("globalnet")));
        assert_eq!(d.threshold_for("marketnet_pet"), (0.05, Some("marketnet")));
        assert_eq!(d.threshold_for("Object_globalnet"), (0.02, Some("globalnet")));
        // whole token only — no substring matches
        assert_eq!(d.threshold_for("globalnetwork"), (0.01, None));
        assert_eq!(d.threshold_for("Object_edu"), (0.01, None));
    }

    #[test]
    fn test_threshold_for_longest_token_wins() {
        let mut config = Config::default();
        config.apply_overrides(&[
            "delta.threshold.net=0.01".into(),
            "delta.threshold.global_net=0.04".into(),
        ]).unwrap();
        // both `net` and `global_net` are whole tokens of `Object_global_net`
        assert_eq!(config.delta.threshold_for("Object_global_net"), (0.04, Some("global_net")));
        assert_eq!(config.delta.threshold_for("market_net"), (0.01, Some("net")));
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
