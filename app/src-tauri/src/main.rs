#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::io::Write;
use std::process::{Command, Stdio};

#[tauri::command]
async fn engine(action: String, payload: serde_json::Value) -> Result<serde_json::Value, String> {
    tauri::async_runtime::spawn_blocking(move || {
        if action != "inspect" && action != "generate" { return Err("Недопустимое действие".into()); }
        let mut command = if cfg!(debug_assertions) {
            let mut cmd = Command::new("python");
            cmd.arg("engine.py");
            cmd.current_dir(env!("CARGO_MANIFEST_DIR").replace("src-tauri", ""));
            cmd
        } else {
            let executable = std::env::current_exe().map_err(|e| e.to_string())?;
            let sidecar = executable.with_file_name("papagun-engine.exe");
            Command::new(sidecar)
        };
        let mut child = command.arg(action).stdin(Stdio::piped()).stdout(Stdio::piped()).stderr(Stdio::piped())
            .spawn().map_err(|e| format!("Не удалось запустить обработку: {e}"))?;
        let input = serde_json::to_vec(&payload).map_err(|e| e.to_string())?;
        child.stdin.take().ok_or("Нет канала передачи данных")?.write_all(&input).map_err(|e| e.to_string())?;
        let result = child.wait_with_output().map_err(|e| e.to_string())?;
        let response: serde_json::Value = serde_json::from_slice(&result.stdout)
            .map_err(|_| String::from_utf8_lossy(&result.stderr).chars().take(500).collect::<String>())?;
        if !result.status.success() { return Err(response["error"].as_str().unwrap_or("Ошибка обработки").to_string()); }
        Ok(response)
    }).await.map_err(|e| e.to_string())?
}

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![engine])
        .run(tauri::generate_context!())
        .expect("Papagun failed to start");
}
