pub mod errors {
    use thiserror::Error;

    #[derive(Debug, Error)]
    pub enum DomainError {
        #[error("validation failed: {0}")]
        Validation(String),
        #[error("serialization failed: {0}")]
        Serialization(String),
    }
}

pub mod modes {
    use serde::{Deserialize, Serialize};

    #[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq, Hash)]
    #[serde(rename_all = "snake_case")]
    pub enum CopilotMode {
        Translation,
        TranslationWithReferences,
        Review,
    }

    impl CopilotMode {
        pub fn label(&self) -> &'static str {
            match self {
                CopilotMode::Translation => "翻訳（通常）",
                CopilotMode::TranslationWithReferences => "翻訳（参照あり）",
                CopilotMode::Review => "翻訳チェック",
            }
        }

        pub fn as_request_value(&self) -> &'static str {
            match self {
                CopilotMode::Translation => "translation",
                CopilotMode::TranslationWithReferences => "translation_with_references",
                CopilotMode::Review => "review",
            }
        }
    }
}

pub mod forms {
    use serde::{Deserialize, Serialize};

    use crate::errors::DomainError;
    use crate::modes::CopilotMode;

    #[derive(Debug, Clone, Serialize, Deserialize, Default)]
    pub struct FormValues {
        pub mode: CopilotMode,
        pub workbook_name: Option<String>,
        pub sheet_name: Option<String>,
        pub cell_range: Option<String>,
        pub translation_output_range: Option<String>,
        pub source_reference_urls: Vec<String>,
        pub target_reference_urls: Vec<String>,
        pub citation_output_range: Option<String>,
        pub source_range: Option<String>,
        pub translated_range: Option<String>,
        pub review_output_range: Option<String>,
    }

    impl FormValues {
        pub fn validate(&self) -> Result<(), DomainError> {
            use CopilotMode::*;

            let missing = |name: &str| DomainError::Validation(format!("{name} は必須です"));

            match self.mode {
                Translation => {
                    self.cell_range.as_ref().ok_or_else(|| missing("cell_range"))?;
                    self.translation_output_range
                        .as_ref()
                        .ok_or_else(|| missing("translation_output_range"))?;
                }
                TranslationWithReferences => {
                    self.cell_range.as_ref().ok_or_else(|| missing("cell_range"))?;
                    self.translation_output_range
                        .as_ref()
                        .ok_or_else(|| missing("translation_output_range"))?;
                    if self.source_reference_urls.is_empty() && self.target_reference_urls.is_empty() {
                        return Err(DomainError::Validation(
                            "参照URLを1つ以上入力してください".into(),
                        ));
                    }
                }
                Review => {
                    self.source_range.as_ref().ok_or_else(|| missing("source_range"))?;
                    self.translated_range
                        .as_ref()
                        .ok_or_else(|| missing("translated_range"))?;
                    self.review_output_range
                        .as_ref()
                        .ok_or_else(|| missing("review_output_range"))?;
                }
            }

            Ok(())
        }
    }
}

pub mod timeline {
    use serde::{Deserialize, Serialize};

    use crate::messages::ResponseType;

    #[derive(Debug, Clone, Serialize, Deserialize)]
    pub struct TimelineEntry {
        pub id: String,
        pub response_type: ResponseType,
        pub content: String,
        pub timestamp: String,
    }
}

pub mod messages {
    use serde::{Deserialize, Serialize};

    use crate::forms::FormValues;

    #[derive(Debug, Clone, Serialize, Deserialize)]
    #[serde(rename_all = "snake_case", tag = "type", content = "data")]
    pub enum RequestPayload {
        UpdateContext { force_refresh: bool },
        SubmitForm(FormValues),
        Stop,
        ResetBrowser,
    }

    #[derive(Debug, Clone, Serialize, Deserialize)]
    pub struct RequestMessage {
        pub request_id: String,
        pub payload: RequestPayload,
    }

    #[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
    #[serde(rename_all = "snake_case")]
    pub enum ResponseType {
        FinalAnswer,
        Error,
        Progress,
        Telemetry,
    }

    #[derive(Debug, Clone, Serialize, Deserialize)]
    #[serde(rename_all = "snake_case", tag = "type", content = "data")]
    pub enum ResponsePayload {
        Progress { message: String, percent: f32 },
        FinalAnswer { content: String },
        Error { message: String },
        Telemetry { content: String },
    }

    #[derive(Debug, Clone, Serialize, Deserialize)]
    pub struct ResponseMessage {
        pub request_id: String,
        pub response_type: ResponseType,
        pub payload: ResponsePayload,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        pub timestamp: Option<String>,
    }
}

pub mod settings {
    use serde::{Deserialize, Serialize};

    use crate::modes::CopilotMode;

    #[derive(Debug, Clone, Serialize, Deserialize, Default)]
    pub struct AppSettings {
        pub last_workbook: Option<String>,
        pub last_sheet: Option<String>,
        pub window_width: Option<f32>,
        pub window_height: Option<f32>,
        pub last_mode: Option<CopilotMode>,
    }
}
