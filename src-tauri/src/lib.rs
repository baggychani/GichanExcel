use std::{
    env,
    path::{Path, PathBuf},
};

use serde::Serialize;
use tauri::{Manager, LogicalSize};
use tauri_plugin_fs::FsExt;

const OPEN_EXTENSIONS: &[&str] = &["xlsx", "xls", "csv", "tsv", "txt"];
const PREFERRED_WINDOW_WIDTH: f64 = 1280.0;
const PREFERRED_WINDOW_HEIGHT: f64 = 860.0;
const MIN_WINDOW_WIDTH: f64 = 960.0;
const MIN_WINDOW_HEIGHT: f64 = 640.0;
const WINDOW_SCREEN_MARGIN: f64 = 40.0;

#[derive(Serialize)]
struct InitialOpenFile {
    path: String,
    bytes: Vec<u8>,
}

fn is_supported_document_extension(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| {
            OPEN_EXTENSIONS
                .iter()
                .any(|supported| extension.eq_ignore_ascii_case(supported))
        })
        .unwrap_or(false)
}

fn is_supported_open_path(path: &Path) -> bool {
    path.is_file() && is_supported_document_extension(path)
}

fn allow_path_in_fs_scope(app: &tauri::AppHandle, path: &Path) -> Result<(), String> {
    if !is_supported_document_extension(path) {
        return Err("지원하지 않는 파일 형식입니다.".into());
    }

    if let Some(scope) = app.try_fs_scope() {
        scope
            .allow_file(path)
            .map_err(|error| error.to_string())?;
    }

    if let Some(scopes) = app.try_state::<tauri::scope::Scopes>() {
        scopes
            .allow_file(path)
            .map_err(|error| error.to_string())?;
    }

    Ok(())
}

fn clamp(value: f64, min: f64, max: f64) -> f64 {
    value.max(min).min(max)
}

/// 창이 보이기 전에 복원 크기·최대화를 끝냅니다.
/// 프론트에서 setSize→maximize 하면 최대화→축소→최대화로 깜빡입니다.
fn prepare_startup_window(window: &tauri::WebviewWindow) -> tauri::Result<()> {
    let monitor = window.current_monitor()?;
    let (restore_width, restore_height) = if let Some(monitor) = monitor {
        let scale = monitor.scale_factor();
        let work_area = monitor.work_area().size;
        let available_width = (work_area.width as f64 / scale) - WINDOW_SCREEN_MARGIN;
        let available_height = (work_area.height as f64 / scale) - WINDOW_SCREEN_MARGIN;
        (
            clamp(PREFERRED_WINDOW_WIDTH, MIN_WINDOW_WIDTH, available_width.max(MIN_WINDOW_WIDTH)),
            clamp(
                PREFERRED_WINDOW_HEIGHT,
                MIN_WINDOW_HEIGHT,
                available_height.max(MIN_WINDOW_HEIGHT),
            ),
        )
    } else {
        (PREFERRED_WINDOW_WIDTH, PREFERRED_WINDOW_HEIGHT)
    };

    window.set_size(LogicalSize::new(restore_width, restore_height))?;
    let _ = window.center();
    window.maximize()?;
    window.show()?;
    let _ = window.set_focus();
    Ok(())
}

#[tauri::command]
fn allow_document_path(app: tauri::AppHandle, path: String) -> Result<(), String> {
    allow_path_in_fs_scope(&app, Path::new(&path))
}

#[tauri::command]
fn initial_open_file(app: tauri::AppHandle) -> Result<Option<InitialOpenFile>, String> {
    let path = env::args_os()
        .skip(1)
        .map(PathBuf::from)
        .find(|path| is_supported_open_path(path));

    let Some(path) = path else {
        return Ok(None);
    };

    allow_path_in_fs_scope(&app, &path)?;
    let bytes = std::fs::read(&path).map_err(|error| error.to_string())?;
    Ok(Some(InitialOpenFile {
        path: path.to_string_lossy().into_owned(),
        bytes,
    }))
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .setup(|app| {
            if let Some(window) = app.get_webview_window("main") {
                // 실패해도 앱은 띄웁니다. show만은 보장합니다.
                if prepare_startup_window(&window).is_err() {
                    let _ = window.maximize();
                    let _ = window.show();
                }
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            initial_open_file,
            allow_document_path
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
