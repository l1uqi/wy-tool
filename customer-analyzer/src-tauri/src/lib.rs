mod excel_processor;
mod monthly_analysis;

use excel_processor::{AnalysisResult, ProcessProgress};
use monthly_analysis::{MonthlyAnalysisResult, CachedRow, CustomerOption};
use std::sync::{Arc, Mutex};
use tauri::{AppHandle, Emitter, State};
use serde::{Deserialize, Serialize};

// 全局状态
struct AppState {
    cancel_flag: Arc<Mutex<bool>>,
    monthly_cache: Arc<Mutex<Option<MonthlyCacheData>>>,
}

#[derive(Clone)]
struct MonthlyCacheData {
    cached_rows: Vec<CachedRow>,
}

/// 加载选项的返回结果
#[derive(Debug, Serialize, Deserialize)]
struct LoadOptionsResult {
    file_path: String,
    available_customers: Vec<CustomerOption>,
    available_provinces: Vec<String>,
    available_cities: Vec<String>,
    available_districts: Vec<String>,
    available_regions: Vec<String>,
    total_rows: usize,
    load_time_ms: u128,
}

#[tauri::command]
async fn analyze_excel(
    file_path: String,
    state: State<'_, AppState>,
    app: AppHandle,
) -> Result<AnalysisResult, String> {
    {
        let mut flag = state.cancel_flag.lock().unwrap();
        *flag = false;
    }

    let cancel_flag = state.cancel_flag.clone();
    let app_handle = app.clone();

    let progress_callback = move |progress: ProcessProgress| {
        let _ = app_handle.emit("excel-progress", &progress);
    };

    let result = tokio::task::spawn_blocking(move || {
        excel_processor::process_excel_file(&file_path, cancel_flag, progress_callback)
    })
    .await
    .map_err(|e| format!("任务执行失败: {}", e))?;

    result
}

/// 加载Excel文件并缓存数据
#[tauri::command]
async fn load_monthly_file(
    file_path: String,
    state: State<'_, AppState>,
    app: AppHandle,
) -> Result<LoadOptionsResult, String> {
    {
        let mut flag = state.cancel_flag.lock().unwrap();
        *flag = false;
    }

    let cancel_flag = state.cancel_flag.clone();
    let app_handle = app.clone();
    let monthly_cache = state.monthly_cache.clone();

    let progress_callback = move |progress: monthly_analysis::ProcessProgress| {
        let _ = app_handle.emit("monthly-progress", &progress);
    };

    let result = tokio::task::spawn_blocking(move || {
        monthly_analysis::load_excel_file(&file_path, cancel_flag, progress_callback)
    })
    .await
    .map_err(|e| format!("任务执行失败: {}", e))??;

    // 缓存数据
    {
        let mut cache = monthly_cache.lock().unwrap();
        *cache = Some(MonthlyCacheData {
            cached_rows: result.cached_rows,
        });
    }

    Ok(LoadOptionsResult {
        file_path: result.file_path,
        available_customers: result.available_customers,
        available_provinces: result.available_provinces,
        available_cities: result.available_cities,
        available_districts: result.available_districts,
        available_regions: result.available_regions,
        total_rows: result.total_rows,
        load_time_ms: result.load_time_ms,
    })
}

/// 基于缓存数据执行月度分析
#[tauri::command]
async fn analyze_monthly_cached(
    analysis_type: String,
    target: String,
    state: State<'_, AppState>,
) -> Result<MonthlyAnalysisResult, String> {
    let monthly_cache = state.monthly_cache.clone();
    
    let result = tokio::task::spawn_blocking(move || {
        let cache = monthly_cache.lock().unwrap();
        
        match cache.as_ref() {
            Some(data) => {
                monthly_analysis::analyze_from_cache(
                    &data.cached_rows, 
                    &analysis_type, 
                    &target
                )
            },
            None => Err("请先加载Excel文件".to_string()),
        }
    })
    .await
    .map_err(|e| format!("任务执行失败: {}", e))??;

    Ok(result)
}

/// 清除缓存
#[tauri::command]
fn clear_monthly_cache(state: State<'_, AppState>) {
    let mut cache = state.monthly_cache.lock().unwrap();
    *cache = None;
}

#[tauri::command]
fn cancel_analysis(state: State<'_, AppState>) {
    let mut flag = state.cancel_flag.lock().unwrap();
    *flag = true;
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_shell::init())
        .manage(AppState {
            cancel_flag: Arc::new(Mutex::new(false)),
            monthly_cache: Arc::new(Mutex::new(None)),
        })
        .invoke_handler(tauri::generate_handler![
            analyze_excel,
            load_monthly_file,
            analyze_monthly_cached,
            clear_monthly_cache,
            cancel_analysis
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
