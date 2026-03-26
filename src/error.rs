use thiserror::Error;

/// All error types for the `oa` CLI.
#[derive(Error, Debug)]
pub enum OaError {
    #[error("COM error: {0}")]
    Com(#[from] windows::core::Error),

    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),

    #[error("ZIP error: {0}")]
    Zip(#[from] zip::result::ZipError),

    #[error("XML error: {0}")]
    Xml(#[from] quick_xml::Error),

    #[error("Config error: {0}")]
    Config(String),

    #[allow(dead_code)]
    #[error("Validation error: {0}")]
    Validation(String),

    #[error("{0}")]
    Other(String),
}

impl From<String> for OaError {
    fn from(s: String) -> Self {
        OaError::Other(s)
    }
}

impl From<&str> for OaError {
    fn from(s: &str) -> Self {
        OaError::Other(s.to_string())
    }
}

pub type OaResult<T> = std::result::Result<T, OaError>;
