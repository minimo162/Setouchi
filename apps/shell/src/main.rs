mod ui;

use std::path::PathBuf;
use std::sync::Arc;

use anyhow::{anyhow, Context, Result};
use serde_json;
use setouchi_copilot::{BrowserChannel, CopilotRunner};
use setouchi_domain::forms::FormValues;
use setouchi_domain::messages::{RequestMessage, RequestPayload, ResponseMessage, ResponsePayload, ResponseType};
use setouchi_excel::{ExcelContext, ExcelService, MockExcelService, PythonExcelService};
use setouchi_telemetry::init_tracing;
use tauri::{async_runtime::Mutex, Manager};
use tokio::sync::{broadcast, mpsc};
use tracing::{error, info, warn};

#[tauri::command]
async fn submit_request(
    payload: RequestMessage,
    tx: tauri::State<'_, AppChannels>,
) -> Result<(), String> {
    tx.request
        .lock()
        .await
        .send(payload)
        .await
        .map_err(|e| e.to_string())
}

struct AppChannels {
    request: Mutex<mpsc::Sender<RequestMessage>>,
    response: broadcast::Sender<ResponseMessage>,
}

struct AppWorker {
    excel: Arc<dyn ExcelService>,
    copilot: Arc<CopilotRunner>,
    responses: broadcast::Sender<ResponseMessage>,
}

impl AppWorker {
    fn new(
        excel: Arc<dyn ExcelService>,
        copilot: Arc<CopilotRunner>,
        responses: broadcast::Sender<ResponseMessage>,
    ) -> Self {
        Self {
            excel,
            copilot,
            responses,
        }
    }

    async fn run(mut self, mut rx: mpsc::Receiver<RequestMessage>) {
        while let Some(message) = rx.recv().await {
            if let Err(err) = self.handle_request(message).await {
                error!(?err, "Failed to process request");
            }
        }
    }

    async fn handle_request(&self, message: RequestMessage) -> Result<()> {
        match message.payload {
            RequestPayload::UpdateContext { .. } => {
                let ctx = self.excel.refresh_context()?;
                self.emit_context(&message.request_id, &ctx)?;
            }
            RequestPayload::SubmitForm(form) => {
                self.process_form(message.request_id, form).await?;
            }
            RequestPayload::Stop => {
                self.emit(ResponseMessage {
                    request_id: message.request_id,
                    response_type: ResponseType::Progress,
                    payload: ResponsePayload::Progress {
                        message: "停止リクエストを受け付けました".into(),
                        percent: 1.0,
                    },
                    timestamp: None,
                })?;
            }
            RequestPayload::ResetBrowser => {
                self.emit(ResponseMessage {
                    request_id: message.request_id,
                    response_type: ResponseType::Telemetry,
                    payload: ResponsePayload::Telemetry {
                        content: "ブラウザリセットを要求しました".into(),
                    },
                    timestamp: None,
                })?;
            }
        }
        Ok(())
    }

    async fn process_form(&self, request_id: String, form: FormValues) -> Result<()> {
        if let Err(err) = form.validate() {
            self.emit(ResponseMessage {
                request_id,
                response_type: ResponseType::Error,
                payload: ResponsePayload::Error {
                    message: err.to_string(),
                },
                timestamp: None,
            })?;
            return Ok(());
        }

        self.copilot.ensure_session().await?;
        let events = self.copilot.run_form(&request_id, &form).await?;
        for event in events {
            self.emit(event)?;
        }
        Ok(())
    }

    fn emit_context(&self, request_id: &str, ctx: &ExcelContext) -> Result<()> {
        let serialized = serde_json::to_string(ctx)?;
        self.emit(ResponseMessage {
            request_id: request_id.to_string(),
            response_type: ResponseType::Telemetry,
            payload: ResponsePayload::Telemetry { content: serialized },
            timestamp: None,
        })
    }

    fn emit(&self, message: ResponseMessage) -> Result<()> {
        self.responses
            .send(message)
            .map(|_| ())
            .map_err(|err| anyhow::anyhow!("failed to broadcast response: {err}"))
    }
}

fn main() -> Result<()> {
    let _ = init_tracing();
    let workspace_root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mock_excel_path = workspace_root.join("assets/mock_excel_context.json");
    let excel_bridge = workspace_root.join("scripts/excel_bridge.py");
    let python_path = std::env::var("SETTOUCHI_PYTHON").unwrap_or_else(|_| "python".into());
    let python_path = PathBuf::from(python_path);

    let excel_service: Arc<dyn ExcelService> = match PythonExcelService::new(python_path.clone(), excel_bridge) {
        Ok(service) => Arc::new(service),
        Err(err) => {
            warn!(?err, "Falling back to mock Excel service");
            Arc::new(MockExcelService::from_file(mock_excel_path))
        }
    };

    let copilot = Arc::new(CopilotRunner::new(
        BrowserChannel::MsEdge,
        python_path.clone(),
        workspace_root.join("scripts/copilot_bridge.py"),
    ));

    let (req_tx, req_rx) = mpsc::channel::<RequestMessage>(32);
    let (res_tx, _) = broadcast::channel::<ResponseMessage>(32);

    tauri::Builder::default()
        .setup(move |app| {
            let handle = app.handle();
            let mut res_rx = res_tx.subscribe();
            tauri::async_runtime::spawn(async move {
                while let Ok(msg) = res_rx.recv().await {
                    if let Err(err) = handle.emit_all("setouchi://response", &msg) {
                        error!(?err, "Failed to emit response event");
                    }
                }
            });

            let worker = AppWorker::new(excel_service.clone(), copilot.clone(), res_tx.clone());
            tauri::async_runtime::spawn(async move {
                worker.run(req_rx).await;
            });

            Ok(())
        })
        .invoke_handler(tauri::generate_handler![submit_request])
        .manage(AppChannels {
            request: Mutex::new(req_tx),
            response: res_tx,
        })
        .run(tauri::generate_context!())?
        .exit(0);

    Ok(())
}
