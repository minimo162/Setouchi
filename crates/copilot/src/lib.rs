#![allow(dead_code)]

use std::path::{Path, PathBuf};

use anyhow::{anyhow, Context, Result};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use tokio::io::{AsyncBufReadExt, BufReader};
use tokio::process::Command;
use tracing::{debug, info, warn};

use setouchi_domain::forms::FormValues;
use setouchi_domain::messages::{ResponseMessage, ResponsePayload, ResponseType};
use setouchi_domain::modes::CopilotMode;

#[derive(Debug, Clone, Serialize)]
pub enum BrowserChannel {
    MsEdge,
    Chromium,
}

pub struct CopilotRunner {
    channel: BrowserChannel,
    python: PathBuf,
    script: PathBuf,
}

impl CopilotRunner {
    pub fn new(channel: BrowserChannel, python: impl Into<PathBuf>, script: impl Into<PathBuf>) -> Self {
        Self {
            channel,
            python: python.into(),
            script: script.into(),
        }
    }

    pub async fn ensure_session(&self) -> Result<()> {
        if !self.python.exists() {
            return Err(anyhow!("python interpreter not found at {:?}", self.python));
        }
        if !self.script.exists() {
            return Err(anyhow!("copilot bridge script not found at {:?}", self.script));
        }
        info!(?self.channel, "Python bridge detected for Copilot");
        Ok(())
    }

    pub async fn run_form(&self, request_id: &str, form: &FormValues) -> Result<Vec<ResponseMessage>> {
        let worker_payload = build_worker_payload(form)?;
        let bridge_input = json!({
            "request_id": request_id,
            "workbook_name": form.workbook_name,
            "sheet_name": form.sheet_name,
            "worker_payload": worker_payload,
        });

        let mut child = Command::new(&self.python)
            .arg(&self.script)
            .arg("run")
            .stdin(std::process::Stdio::piped())
            .stdout(std::process::Stdio::piped())
            .stderr(std::process::Stdio::piped())
            .spawn()
            .with_context(|| "failed to spawn copilot bridge")?;

        if let Some(mut stdin) = child.stdin.take() {
            let payload_bytes = serde_json::to_vec(&bridge_input)?;
            tokio::spawn(async move {
                use tokio::io::AsyncWriteExt;
                if let Err(err) = stdin.write_all(&payload_bytes).await {
                    warn!(?err, "failed to send payload to copilot bridge");
                }
            });
        }

        let stdout = child.stdout.take().ok_or_else(|| anyhow!("bridge stdout unavailable"))?;
        let mut reader = BufReader::new(stdout).lines();
        let mut responses = Vec::new();

        while let Some(line) = reader.next_line().await? {
            let trimmed = line.trim();
            if trimmed.is_empty() {
                continue;
            }
            match serde_json::from_str::<BridgeEvent>(trimmed) {
                Ok(event) => {
                    if let Some(message) = event.into_message(request_id) {
                        responses.push(message);
                    }
                }
                Err(err) => {
                    debug!(?err, line = trimmed, "Ignoring non-JSON bridge output");
                }
            }
        }

        let status = child.wait().await?;
        if !status.success() && responses.is_empty() {
            return Err(anyhow!("copilot bridge exited with status {status}"));
        }

        Ok(responses)
    }
}

#[derive(Debug, Deserialize)]
struct BridgeEvent {
    event: String,
    #[serde(rename = "type")]
    kind: Option<String>,
    content: Option<String>,
    metadata: Option<Value>,
    message: Option<String>,
}

impl BridgeEvent {
    fn into_message(self, request_id: &str) -> Option<ResponseMessage> {
        if self.event != "response" {
            return None;
        }
        let response_type = match self.kind.as_deref().unwrap_or("") {
            "progress" => ResponseType::Progress,
            "error" => ResponseType::Error,
            "final_answer" => ResponseType::FinalAnswer,
            "status" => ResponseType::Telemetry,
            other => {
                warn!(kind = other, "Unknown response type from bridge");
                ResponseType::Telemetry
            }
        };

        let payload = match response_type {
            ResponseType::Progress => ResponsePayload::Progress {
                message: self.content.or(self.message).unwrap_or_default(),
                percent: self
                    .metadata
                    .as_ref()
                    .and_then(|meta| meta.get("percent"))
                    .and_then(|value| value.as_f64())
                    .unwrap_or(0.0) as f32,
            },
            ResponseType::Error => ResponsePayload::Error {
                message: self.content.or(self.message).unwrap_or_default(),
            },
            ResponseType::FinalAnswer => ResponsePayload::FinalAnswer {
                content: self.content.unwrap_or_default(),
            },
            ResponseType::Telemetry => ResponsePayload::Telemetry {
                content: self.content.unwrap_or_else(|| "telemetry".into()),
            },
            other => {
                warn!(?other, "Unhandled response type");
                ResponsePayload::Telemetry {
                    content: self.content.unwrap_or_default(),
                }
            }
        };

        Some(ResponseMessage {
            request_id: request_id.to_string(),
            response_type,
            payload,
            timestamp: None,
        })
    }
}

fn build_worker_payload(form: &FormValues) -> Result<Value> {
    let mut arguments = serde_json::Map::new();
    match form.mode {
        CopilotMode::Translation => {
            arguments.insert("cell_range".into(), require_string(&form.cell_range, "cell_range")?.into());
            arguments.insert(
                "translation_output_range".into(),
                require_string(&form.translation_output_range, "translation_output_range")?.into(),
            );
            arguments.insert("target_language".into(), Value::String("English".into()));
        }
        CopilotMode::TranslationWithReferences => {
            arguments.insert("cell_range".into(), require_string(&form.cell_range, "cell_range")?.into());
            arguments.insert(
                "translation_output_range".into(),
                require_string(&form.translation_output_range, "translation_output_range")?.into(),
            );
            if form.source_reference_urls.is_empty() && form.target_reference_urls.is_empty() {
                return Err(anyhow!("参照URLを1件以上入力してください。"));
            }
            if !form.source_reference_urls.is_empty() {
                arguments.insert(
                    "source_reference_urls".into(),
                    Value::Array(form.source_reference_urls.iter().map(|v| Value::String(v.clone())).collect()),
                );
            }
            if !form.target_reference_urls.is_empty() {
                arguments.insert(
                    "target_reference_urls".into(),
                    Value::Array(form.target_reference_urls.iter().map(|v| Value::String(v.clone())).collect()),
                );
            }
            arguments.insert("target_language".into(), Value::String("English".into()));
        }
        CopilotMode::Review => {
            arguments.insert("source_range".into(), require_string(&form.source_range, "source_range")?.into());
            arguments.insert(
                "translated_range".into(),
                require_string(&form.translated_range, "translated_range")?.into(),
            );
            let combined = require_string(&form.review_output_range, "review_output_range")?;
            for (key, range) in derive_review_ranges(&combined)? {
                arguments.insert(key, Value::String(range));
            }
        }
    }

    let tool_name = match form.mode {
        CopilotMode::Translation => "translate_range_without_references",
        CopilotMode::TranslationWithReferences => "translate_range_with_references",
        CopilotMode::Review => "check_translation_quality",
    };

    let mut payload = serde_json::Map::new();
    payload.insert("mode".into(), Value::String(form.mode.as_request_value().into()));
    payload.insert("tool_name".into(), Value::String(tool_name.into()));
    payload.insert("arguments".into(), Value::Object(arguments));
    if let Some(book) = &form.workbook_name {
        payload.insert("workbook_name".into(), Value::String(book.clone()));
    }
    if let Some(sheet) = &form.sheet_name {
        payload.insert("sheet_name".into(), Value::String(sheet.clone()));
    }

    Ok(Value::Object(payload))
}

fn require_string(value: &Option<String>, field: &str) -> Result<String> {
    value
        .as_ref()
        .map(|v| v.trim().to_string())
        .filter(|v| !v.is_empty())
        .ok_or_else(|| anyhow!(format!("{field} を入力してください。")))
}

fn derive_review_ranges(combined: &str) -> Result<Vec<(String, String)>> {
    let (sheet, range) = split_sheet_and_range(combined);
    let (start, end) = range
        .split_once(':')
        .ok_or_else(|| anyhow!("出力範囲は開始セルと終了セルを『:』で区切ってください。"))?;
    let (start_col, start_row) = parse_cell_reference(start)?;
    let (end_col, end_row) = parse_cell_reference(end)?;
    if end_row < start_row {
        return Err(anyhow!("出力範囲の終了行は開始行以上にしてください。"));
    }
    let start_idx = column_label_to_index(&start_col)?;
    let end_idx = column_label_to_index(&end_col)?;
    if end_idx < start_idx {
        return Err(anyhow!("出力範囲の列指定が逆転しています。"));
    }
    let column_count = end_idx - start_idx + 1;
    if column_count < 3 {
        return Err(anyhow!("出力範囲には少なくとも3列が必要です。"));
    }
    if column_count > 4 {
        return Err(anyhow!("出力範囲は最大4列までです。"));
    }
    let mut results = Vec::new();
    let prefix = sheet.map(|s| format!("{s}!")).unwrap_or_default();
    let labels: [(&str, u32); 4] = [
        ("status_output_range", 0),
        ("issue_output_range", 1),
        ("highlight_output_range", 2),
        ("corrected_output_range", 3),
    ];
    for (name, offset) in labels.iter().copied() {
        if offset >= column_count {
            break;
        }
        let col_label = column_index_to_label(start_idx + offset)?;
        let value = format!("{prefix}{col}{start_row}:{col}{end_row}", col = col_label);
        results.push((name.into(), value));
    }
    Ok(results)
}

fn split_sheet_and_range(value: &str) -> (Option<String>, String) {
    if let Some(pos) = value.rfind('!') {
        let (sheet, range) = value.split_at(pos);
        let mut sheet = sheet.trim().to_string();
        if sheet.starts_with('\'') && sheet.ends_with('\'') && sheet.len() >= 2 {
            sheet = sheet[1..sheet.len() - 1].to_string();
        }
        let range = range.trim_start_matches('!').to_string();
        (Some(sheet), range)
    } else {
        (None, value.to_string())
    }
}

fn parse_cell_reference(value: &str) -> Result<(String, u32)> {
    let mut letters = String::new();
    let mut digits = String::new();
    for ch in value.chars() {
        if ch.is_ascii_alphabetic() {
            letters.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            digits.push(ch);
        }
    }
    if letters.is_empty() || digits.is_empty() {
        return Err(anyhow!("セル参照の書式が正しくありません: {value}"));
    }
    let row: u32 = digits.parse()?;
    Ok((letters, row))
}

fn column_label_to_index(label: &str) -> Result<u32> {
    let mut value = 0u32;
    for ch in label.chars() {
        if !ch.is_ascii_alphabetic() {
            return Err(anyhow!("列名に無効な文字が含まれています: {label}"));
        }
        value = value * 26 + ((ch.to_ascii_uppercase() as u8 - b'A') as u32 + 1);
    }
    Ok(value)
}

fn column_index_to_label(mut index: u32) -> Result<String> {
    if index == 0 {
        return Err(anyhow!("列インデックスは1以上で指定してください"));
    }
    let mut label = String::new();
    while index > 0 {
        let rem = ((index - 1) % 26) as u8;
        label.insert(0, (b'A' + rem) as char);
        index = (index - 1) / 26;
    }
    Ok(label)
}
