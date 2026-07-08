use std::{
    env,
    path::{Path, PathBuf},
};

use serde::Serialize;

const OPEN_EXTENSIONS: &[&str] = &["xlsx", "xls", "csv", "tsv", "txt"];

#[derive(Serialize)]
struct InitialOpenFile {
    path: String,
    bytes: Vec<u8>,
}

fn is_supported_open_path(path: &Path) -> bool {
    path.is_file()
        && path
            .extension()
            .and_then(|extension| extension.to_str())
            .map(|extension| {
                OPEN_EXTENSIONS
                    .iter()
                    .any(|supported| extension.eq_ignore_ascii_case(supported))
            })
            .unwrap_or(false)
}

#[tauri::command]
fn initial_open_file() -> Result<Option<InitialOpenFile>, String> {
    let path = env::args_os()
        .skip(1)
        .map(PathBuf::from)
        .find(|path| is_supported_open_path(path));

    let Some(path) = path else {
        return Ok(None);
    };

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
        .invoke_handler(tauri::generate_handler![initial_open_file])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
