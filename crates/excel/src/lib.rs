#![allow(dead_code)]

use std::{fs, path::PathBuf, process::Command, sync::RwLock};

use anyhow::{anyhow, Context, Result};
use serde::{Deserialize, Serialize};
use serde_json;
use tracing::{debug, info, warn};

pub trait ExcelService: Send + Sync {
    fn refresh_context(&self) -> Result<ExcelContext>;
    fn focus_workbook(&self, workbook: &str) -> Result<()>;
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ExcelContext {
    pub active_workbook: Option<String>,
    pub active_sheet: Option<String>,
    pub workbooks: Vec<WorkbookInfo>,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct WorkbookInfo {
    pub name: String,
    pub sheets: Vec<String>,
}

pub struct MockExcelService {
    source: PathBuf,
    cache: RwLock<Option<ExcelContext>>,
}

impl MockExcelService {
    pub fn from_file(path: impl Into<PathBuf>) -> Self {
        Self {
            source: path.into(),
            cache: RwLock::new(None),
        }
    }

    fn load_from_disk(&self) -> Result<ExcelContext> {
        let raw = fs::read_to_string(&self.source)
            .with_context(|| format!("Failed to read mock excel data at {:?}", self.source))?;
        let ctx: ExcelContext = serde_json::from_str(&raw)
            .with_context(|| "Failed to parse mock excel context JSON".to_string())?;
        Ok(ctx)
    }
}

impl ExcelService for MockExcelService {
    fn refresh_context(&self) -> Result<ExcelContext> {
        let ctx = self.load_from_disk()?;
        *self.cache.write().expect("excel cache poisoned") = Some(ctx.clone());
        info!(workbooks = ctx.workbooks.len(), "Mock Excel context refreshed");
        Ok(ctx)
    }

    fn focus_workbook(&self, workbook: &str) -> Result<()> {
        debug!(%workbook, "mock focus workbook");
        Ok(())
    }
}

pub struct PythonExcelService {
    python: PathBuf,
    script: PathBuf,
}

impl PythonExcelService {
    pub fn new(python: impl Into<PathBuf>, script: impl Into<PathBuf>) -> Result<Self> {
        let python = python.into();
        let script = script.into();
        if !python.exists() {
            return Err(anyhow!("python interpreter not found at {:?}", python));
        }
        if !script.exists() {
            return Err(anyhow!("excel bridge script not found at {:?}", script));
        }
        Ok(Self { python, script })
    }

    fn invoke_json(&self, args: &[&str]) -> Result<serde_json::Value> {
        let output = Command::new(&self.python)
            .arg(&self.script)
            .args(args)
            .output()
            .with_context(|| "failed to launch excel bridge script")?;
        if !output.status.success() {
            warn!("excel_bridge_exit" = ?output.status, stderr = %String::from_utf8_lossy(&output.stderr));
        }
        let stdout = String::from_utf8_lossy(&output.stdout);
        let last_line = stdout.lines().filter(|line| !line.trim().is_empty()).last().ok_or_else(|| {
            anyhow!("excel bridge produced no output. stderr={}", String::from_utf8_lossy(&output.stderr))
        })?;
        let value: serde_json::Value = serde_json::from_str(last_line)
            .with_context(|| format!("failed to parse excel bridge output: {last_line}"))?;
        if let Some(error) = value.get("error").and_then(|v| v.as_str()) {
            return Err(anyhow!("excel bridge error: {error}"));
        }
        Ok(value)
    }
}

impl ExcelService for PythonExcelService {
    fn refresh_context(&self) -> Result<ExcelContext> {
        let value = self.invoke_json(&["refresh-context"])?;
        let ctx_value = value
            .get("context")
            .cloned()
            .ok_or_else(|| anyhow!("excel bridge did not return context payload"))?;
        let ctx: ExcelContext = serde_json::from_value(ctx_value)?;
        Ok(ctx)
    }

    fn focus_workbook(&self, workbook: &str) -> Result<()> {
        let _ = self.invoke_json(&["focus-workbook", workbook])?;
        Ok(())
    }
}

#[cfg(target_os = "windows")]
pub mod interop {
    use anyhow::{Context, Result};
    use tracing::info;
    use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};

    pub struct ApartmentToken;

    impl ApartmentToken {
        pub fn new() -> Result<Self> {
            unsafe {
                CoInitializeEx(None, COINIT_APARTMENTTHREADED)
                    .context("Failed to initialize COM apartment")?;
            }
            info!("excel", "COM apartment initialized");
            Ok(Self)
        }
    }

    impl Drop for ApartmentToken {
        fn drop(&mut self) {
            unsafe {
                CoUninitialize();
            }
        }
    }
}
