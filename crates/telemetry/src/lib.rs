use anyhow::Result;
use directories::ProjectDirs;
use std::path::PathBuf;
use tracing_appender::rolling::{RollingFileAppender, Rotation};
use tracing_subscriber::{fmt, EnvFilter};

pub fn init_tracing() -> Result<PathBuf> {
    let dirs = ProjectDirs::from("com", "Setouchi", "Setouchi")
        .expect("project dirs");
    let log_dir = dirs.data_dir().join("logs");
    std::fs::create_dir_all(&log_dir)?;
    let file_appender = RollingFileAppender::new(Rotation::DAILY, &log_dir, "app.log");
    let filter = EnvFilter::try_from_default_env().unwrap_or_else(|_| EnvFilter::new("info"));
    fmt()
        .with_env_filter(filter)
        .with_writer(file_appender)
        .try_init()
        .ok();
    Ok(log_dir)
}
