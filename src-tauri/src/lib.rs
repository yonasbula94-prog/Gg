use serde::Serialize;

#[derive(Serialize)]
struct DesktopRuntimeInfo {
    platform: &'static str,
    version: &'static str,
    offline_cache_supported: bool,
}

#[tauri::command]
fn get_desktop_runtime_info() -> DesktopRuntimeInfo {
    DesktopRuntimeInfo {
        platform: std::env::consts::OS,
        version: env!("CARGO_PKG_VERSION"),
        offline_cache_supported: true,
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![get_desktop_runtime_info])
        .run(tauri::generate_context!())
        .expect("error while running Grand Group Procurement Suite desktop client");
}
