use std::{
    env,
    path::{Path, PathBuf},
};

use serde::Serialize;
use tauri::Manager;
use tauri_plugin_fs::FsExt;

const OPEN_EXTENSIONS: &[&str] = &["xlsx", "xls", "csv", "tsv", "txt"];

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
        .invoke_handler(tauri::generate_handler![
            initial_open_file,
            allow_document_path
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
