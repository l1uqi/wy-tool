mod excel_processor;
mod monthly_analysis;

use excel_processor::{AnalysisResult, ProcessProgress, CustomerData};
use monthly_analysis::{MonthlyAnalysisResult, CachedRow, CustomerOption};
use std::sync::{Arc, Mutex};
use tauri::{AppHandle, Emitter, State};
use serde::{Deserialize, Serialize};
use std::path::PathBuf;
use std::fs;
use std::env;

// 全局状态
struct AppState {
    cancel_flag: Arc<Mutex<bool>>,
    data_cache: Arc<Mutex<Option<DataCache>>>,
}

#[derive(Clone)]
struct DataCache {
    file_path: String,
    cached_rows: Vec<CachedRow>,
    // 用于前20大客户分析的缓存
    customer_data_map: std::collections::HashMap<String, CustomerData>,
}

/// 数据源配置
#[derive(Debug, Serialize, Deserialize)]
struct DataSourceConfig {
    file_path: String,
    file_name: String,
    loaded_at: String,
}

/// 加载选项的返回结果
#[derive(Debug, Serialize, Deserialize)]
struct LoadOptionsResult {
    file_path: String,
    file_name: String,
    available_customers: Vec<CustomerOption>,
    available_provinces: Vec<String>,
    available_cities: Vec<String>,
    available_districts: Vec<String>,
    available_regions: Vec<String>,
    total_rows: usize,
    load_time_ms: u128,
}

/// 数据源信息
#[derive(Debug, Serialize, Deserialize)]
struct DataSourceInfo {
    file_path: String,
    file_name: String,
    loaded_at: String,
    total_rows: usize,
}

/// 获取配置文件路径
fn get_config_path(_app: &AppHandle) -> PathBuf {
    // 使用用户数据目录
    let app_data_dir = dirs::data_dir()
        .map(|d| d.join("customer-analyzer"))
        .unwrap_or_else(|| {
            // 如果获取用户数据目录失败，使用当前目录
            env::current_dir().unwrap_or_else(|_| PathBuf::from("."))
        });
    
    std::fs::create_dir_all(&app_data_dir).unwrap_or_default();
    app_data_dir.join("data_source.json")
}

/// 保存数据源配置
fn save_data_source_config(app: &AppHandle, file_path: &str) -> Result<(), String> {
    let config_path = get_config_path(app);
    let file_name = std::path::Path::new(file_path)
        .file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("未知文件")
        .to_string();
    
    let config = DataSourceConfig {
        file_path: file_path.to_string(),
        file_name,
        loaded_at: chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string(),
    };
    
    let json = serde_json::to_string_pretty(&config)
        .map_err(|e| format!("序列化配置失败: {}", e))?;
    
    fs::write(&config_path, json)
        .map_err(|e| format!("保存配置失败: {}", e))?;
    
    Ok(())
}

/// 读取数据源配置
fn load_data_source_config(app: &AppHandle) -> Result<Option<DataSourceConfig>, String> {
    let config_path = get_config_path(app);
    
    if !config_path.exists() {
        return Ok(None);
    }
    
    let content = fs::read_to_string(&config_path)
        .map_err(|e| format!("读取配置失败: {}", e))?;
    
    let config: DataSourceConfig = serde_json::from_str(&content)
        .map_err(|e| format!("解析配置失败: {}", e))?;
    
    // 检查文件是否还存在
    if !std::path::Path::new(&config.file_path).exists() {
        return Ok(None);
    }
    
    Ok(Some(config))
}

/// 设置数据源（从首页导入）
#[tauri::command]
async fn set_data_source(
    file_path: String,
    state: State<'_, AppState>,
    app: AppHandle,
) -> Result<LoadOptionsResult, String> {
    {
        let mut flag = state.cancel_flag.lock().unwrap();
        *flag = false;
    }

    // 保存配置
    save_data_source_config(&app, &file_path)?;

    // 加载并缓存数据
    let cancel_flag = state.cancel_flag.clone();
    let app_handle = app.clone();
    let data_cache = state.data_cache.clone();

    let progress_callback = move |progress: monthly_analysis::ProcessProgress| {
        let _ = app_handle.emit("excel-progress", ProcessProgress {
            step: progress.step,
            message: progress.message,
            percent: progress.percent,
            detail: progress.detail,
        });
    };

    let result = tokio::task::spawn_blocking({
        let file_path = file_path.clone();
        move || {
            monthly_analysis::load_excel_file(&file_path, cancel_flag, progress_callback)
        }
    })
    .await
    .map_err(|e| format!("任务执行失败: {}", e))??;

    // 构建客户数据映射（用于前20大客户分析）
    let mut customer_data_map: std::collections::HashMap<String, CustomerData> = 
        std::collections::HashMap::new();
    
    for row in &result.cached_rows {
        customer_data_map
            .entry(row.customer_code.clone())
            .and_modify(|existing| {
                existing.pay_amount += row.pay_amount;
                existing.recharge_deduction += row.recharge_deduction;
                existing.total_amount += row.total_amount;
                existing.order_count += 1;
                if existing.customer_name.is_empty() && !row.customer_name.is_empty() {
                    existing.customer_name = row.customer_name.clone();
                }
            })
            .or_insert_with(|| CustomerData {
                customer_code: row.customer_code.clone(),
                customer_name: row.customer_name.clone(),
                pay_amount: row.pay_amount,
                recharge_deduction: row.recharge_deduction,
                total_amount: row.total_amount,
                order_count: 1,
            });
    }

    // 缓存数据
    {
        let mut cache = data_cache.lock().unwrap();
        *cache = Some(DataCache {
            file_path: result.file_path.clone(),
            cached_rows: result.cached_rows.clone(),
            customer_data_map,
        });
    }

    let file_name = std::path::Path::new(&result.file_path)
        .file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("未知文件")
        .to_string();

    Ok(LoadOptionsResult {
        file_path: result.file_path,
        file_name,
        available_customers: result.available_customers,
        available_provinces: result.available_provinces,
        available_cities: result.available_cities,
        available_districts: result.available_districts,
        available_regions: result.available_regions,
        total_rows: result.total_rows,
        load_time_ms: result.load_time_ms,
    })
}

/// 获取当前数据源信息
#[tauri::command]
async fn get_data_source_info(
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<Option<DataSourceInfo>, String> {
    // 先尝试从配置读取
    if let Some(config) = load_data_source_config(&app)? {
        // 检查缓存是否存在
        let cache = state.data_cache.lock().unwrap();
        if let Some(ref data_cache) = *cache {
            if data_cache.file_path == config.file_path {
                return Ok(Some(DataSourceInfo {
                    file_path: config.file_path.clone(),
                    file_name: config.file_name.clone(),
                    loaded_at: config.loaded_at,
                    total_rows: data_cache.cached_rows.len(),
                }));
            }
        }
        
        // 如果缓存不存在但配置存在，尝试加载
        return Ok(Some(DataSourceInfo {
            file_path: config.file_path.clone(),
            file_name: config.file_name.clone(),
            loaded_at: config.loaded_at,
            total_rows: 0, // 需要重新加载
        }));
    }
    
    Ok(None)
}

/// 自动加载数据源（如果存在）
#[tauri::command]
async fn auto_load_data_source(
    state: State<'_, AppState>,
    app: AppHandle,
) -> Result<Option<LoadOptionsResult>, String> {
    let config = match load_data_source_config(&app)? {
        Some(c) => c,
        None => return Ok(None),
    };

    // 如果缓存已存在且文件路径匹配，直接返回
    {
        let cache = state.data_cache.lock().unwrap();
        if let Some(ref data_cache) = *cache {
            if data_cache.file_path == config.file_path {
                // 从缓存构建返回结果
                let mut customers_map: std::collections::HashMap<String, String> = 
                    std::collections::HashMap::new();
                let mut provinces_set: std::collections::HashSet<String> = 
                    std::collections::HashSet::new();
                let mut cities_set: std::collections::HashSet<String> = 
                    std::collections::HashSet::new();
                let mut districts_set: std::collections::HashSet<String> = 
                    std::collections::HashSet::new();
                let mut regions_set: std::collections::HashSet<String> = 
                    std::collections::HashSet::new();

                for row in &data_cache.cached_rows {
                    if !row.customer_code.is_empty() {
                        customers_map.insert(row.customer_code.clone(), row.customer_name.clone());
                    }
                    if let Some(ref p) = row.province {
                        if !p.is_empty() { provinces_set.insert(p.clone()); }
                    }
                    if let Some(ref c) = row.city {
                        if !c.is_empty() { cities_set.insert(c.clone()); }
                    }
                    if let Some(ref d) = row.district {
                        if !d.is_empty() { districts_set.insert(d.clone()); }
                    }
                    if let Some(ref r) = row.region {
                        if !r.is_empty() { regions_set.insert(r.clone()); }
                    }
                }

                let available_customers: Vec<CustomerOption> = customers_map
                    .into_iter()
                    .map(|(code, name)| CustomerOption { code, name })
                    .collect();

                let mut available_provinces: Vec<String> = provinces_set.into_iter().collect();
                available_provinces.sort();
                let mut available_cities: Vec<String> = cities_set.into_iter().collect();
                available_cities.sort();
                let mut available_districts: Vec<String> = districts_set.into_iter().collect();
                available_districts.sort();
                let mut available_regions: Vec<String> = regions_set.into_iter().collect();
                available_regions.sort();

                return Ok(Some(LoadOptionsResult {
                    file_path: config.file_path.clone(),
                    file_name: config.file_name.clone(),
                    available_customers,
                    available_provinces,
                    available_cities,
                    available_districts,
                    available_regions,
                    total_rows: data_cache.cached_rows.len(),
                    load_time_ms: 0,
                }));
            }
        }
    }

    // 缓存不存在，需要重新加载
    let result = set_data_source(config.file_path.clone(), state, app).await?;
    Ok(Some(result))
}

/// 前20大客户分析（使用缓存数据）
#[tauri::command]
async fn analyze_top20_cached(
    state: State<'_, AppState>,
) -> Result<AnalysisResult, String> {
    let data_cache = state.data_cache.clone();
    
    let result = tokio::task::spawn_blocking(move || {
        let cache = data_cache.lock().unwrap();
        
        match cache.as_ref() {
            Some(data) => {
                let mut customers: Vec<CustomerData> = 
                    data.customer_data_map.values().cloned().collect();
                
                // 排序
                customers.sort_by(|a, b| {
                    b.total_amount
                        .partial_cmp(&a.total_amount)
                        .unwrap_or(std::cmp::Ordering::Equal)
                });
                
                let total_amount: f64 = customers.iter().map(|c| c.total_amount).sum();
                let top20: Vec<CustomerData> = customers.iter().take(20).cloned().collect();
                let top20_amount: f64 = top20.iter().map(|c| c.total_amount).sum();
                
                Ok(AnalysisResult {
                    top20,
                    total_customers: customers.len(),
                    total_amount,
                    top20_amount,
                    total_rows: data.cached_rows.len(),
                    process_time_ms: 0,
                })
            },
            None => Err("请先在首页导入数据源".to_string()),
        }
    })
    .await
    .map_err(|e| format!("任务执行失败: {}", e))??;

    Ok(result)
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

/// 加载Excel文件并缓存数据（保留兼容性）
#[tauri::command]
async fn load_monthly_file(
    file_path: String,
    state: State<'_, AppState>,
    app: AppHandle,
) -> Result<LoadOptionsResult, String> {
    set_data_source(file_path, state, app).await
}

/// 基于缓存数据执行月度分析
#[tauri::command]
async fn analyze_monthly_cached(
    analysis_type: String,
    target: String,
    state: State<'_, AppState>,
) -> Result<MonthlyAnalysisResult, String> {
    let data_cache = state.data_cache.clone();
    
    let result = tokio::task::spawn_blocking(move || {
        let cache = data_cache.lock().unwrap();
        
        match cache.as_ref() {
            Some(data) => {
                monthly_analysis::analyze_from_cache(
                    &data.cached_rows, 
                    &analysis_type, 
                    &target
                )
            },
            None => Err("请先在首页导入数据源".to_string()),
        }
    })
    .await
    .map_err(|e| format!("任务执行失败: {}", e))??;

    Ok(result)
}

/// 获取月度分析的选项（从缓存）
#[tauri::command]
async fn get_monthly_options(
    state: State<'_, AppState>,
) -> Result<Option<LoadOptionsResult>, String> {
    let data_cache = state.data_cache.clone();
    
    let result: Option<LoadOptionsResult> = tokio::task::spawn_blocking(move || {
        let cache = data_cache.lock().unwrap();
        
        match cache.as_ref() {
            Some(data) => {
                let mut customers_map: std::collections::HashMap<String, String> = 
                    std::collections::HashMap::new();
                let mut provinces_set: std::collections::HashSet<String> = 
                    std::collections::HashSet::new();
                let mut cities_set: std::collections::HashSet<String> = 
                    std::collections::HashSet::new();
                let mut districts_set: std::collections::HashSet<String> = 
                    std::collections::HashSet::new();
                let mut regions_set: std::collections::HashSet<String> = 
                    std::collections::HashSet::new();

                for row in &data.cached_rows {
                    if !row.customer_code.is_empty() {
                        customers_map.insert(row.customer_code.clone(), row.customer_name.clone());
                    }
                    if let Some(ref p) = row.province {
                        if !p.is_empty() { provinces_set.insert(p.clone()); }
                    }
                    if let Some(ref c) = row.city {
                        if !c.is_empty() { cities_set.insert(c.clone()); }
                    }
                    if let Some(ref d) = row.district {
                        if !d.is_empty() { districts_set.insert(d.clone()); }
                    }
                    if let Some(ref r) = row.region {
                        if !r.is_empty() { regions_set.insert(r.clone()); }
                    }
                }

                let available_customers: Vec<CustomerOption> = customers_map
                    .into_iter()
                    .map(|(code, name)| CustomerOption { code, name })
                    .collect();

                let mut available_provinces: Vec<String> = provinces_set.into_iter().collect();
                available_provinces.sort();
                let mut available_cities: Vec<String> = cities_set.into_iter().collect();
                available_cities.sort();
                let mut available_districts: Vec<String> = districts_set.into_iter().collect();
                available_districts.sort();
                let mut available_regions: Vec<String> = regions_set.into_iter().collect();
                available_regions.sort();

                let file_name = std::path::Path::new(&data.file_path)
                    .file_name()
                    .and_then(|n| n.to_str())
                    .unwrap_or("未知文件")
                    .to_string();

                Some(LoadOptionsResult {
                    file_path: data.file_path.clone(),
                    file_name,
                    available_customers,
                    available_provinces,
                    available_cities,
                    available_districts,
                    available_regions,
                    total_rows: data.cached_rows.len(),
                    load_time_ms: 0,
                })
            },
            None => None,
        }
    })
    .await
    .map_err(|e| format!("任务执行失败: {}", e))?;

    Ok(result)
}

/// 清除缓存
#[tauri::command]
fn clear_data_cache(state: State<'_, AppState>) {
    let mut cache = state.data_cache.lock().unwrap();
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
            data_cache: Arc::new(Mutex::new(None)),
        })
        .invoke_handler(tauri::generate_handler![
            analyze_excel,
            set_data_source,
            get_data_source_info,
            auto_load_data_source,
            analyze_top20_cached,
            load_monthly_file,
            get_monthly_options,
            analyze_monthly_cached,
            clear_data_cache,
            cancel_analysis
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
