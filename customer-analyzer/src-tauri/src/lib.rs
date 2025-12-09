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

/// 单个数据源配置
#[derive(Debug, Serialize, Deserialize, Clone)]
struct DataSourceConfig {
    id: String,
    file_path: String,
    file_name: String,
    loaded_at: String,
    total_rows: usize,
}

/// 数据源列表配置
#[derive(Debug, Serialize, Deserialize)]
struct DataSourceListConfig {
    data_sources: Vec<DataSourceConfig>,
    current_id: Option<String>,
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

/// 数据源信息（单个）
#[derive(Debug, Serialize, Deserialize, Clone)]
struct DataSourceInfo {
    id: String,
    file_path: String,
    file_name: String,
    loaded_at: String,
    total_rows: usize,
}

/// 数据源列表信息
#[derive(Debug, Serialize, Deserialize)]
struct DataSourceListInfo {
    data_sources: Vec<DataSourceInfo>,
    current_id: Option<String>,
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

/// 读取数据源列表配置
fn load_data_source_list_config(app: &AppHandle) -> Result<DataSourceListConfig, String> {
    let config_path = get_config_path(app);
    
    if !config_path.exists() {
        return Ok(DataSourceListConfig {
            data_sources: Vec::new(),
            current_id: None,
        });
    }
    
    let content = fs::read_to_string(&config_path)
        .map_err(|e| format!("读取配置失败: {}", e))?;
    
    // 尝试解析为新格式
    match serde_json::from_str::<DataSourceListConfig>(&content) {
        Ok(mut config) => {
            // 过滤掉文件不存在的数据源
            config.data_sources.retain(|ds| std::path::Path::new(&ds.file_path).exists());
            Ok(config)
        },
        Err(_) => {
            // 尝试解析为旧格式（单个数据源）
            match serde_json::from_str::<serde_json::Value>(&content) {
                Ok(old_value) => {
                    if let (Some(file_path), Some(file_name)) = (
                        old_value.get("file_path").and_then(|v| v.as_str()),
                        old_value.get("file_name").and_then(|v| v.as_str()),
                    ) {
                        if std::path::Path::new(file_path).exists() {
                            let id = uuid::Uuid::new_v4().to_string();
                            Ok(DataSourceListConfig {
                                data_sources: vec![DataSourceConfig {
                                    id: id.clone(),
                                    file_path: file_path.to_string(),
                                    file_name: file_name.to_string(),
                                    loaded_at: old_value.get("loaded_at")
                                        .and_then(|v| v.as_str())
                                        .unwrap_or("")
                                        .to_string(),
                                    total_rows: 0,
                                }],
                                current_id: Some(id),
                            })
                        } else {
                            Ok(DataSourceListConfig {
                                data_sources: Vec::new(),
                                current_id: None,
                            })
                        }
                    } else {
                        Ok(DataSourceListConfig {
                            data_sources: Vec::new(),
                            current_id: None,
                        })
                    }
                },
                Err(_) => Ok(DataSourceListConfig {
                    data_sources: Vec::new(),
                    current_id: None,
                }),
            }
        },
    }
}

/// 保存数据源列表配置
fn save_data_source_list_config(app: &AppHandle, config: &DataSourceListConfig) -> Result<(), String> {
    let config_path = get_config_path(app);
    
    let json = serde_json::to_string_pretty(config)
        .map_err(|e| format!("序列化配置失败: {}", e))?;
    
    fs::write(&config_path, json)
        .map_err(|e| format!("保存配置失败: {}", e))?;
    
    Ok(())
}

/// 添加数据源（从首页导入）
#[tauri::command]
async fn add_data_source(
    file_path: String,
    state: State<'_, AppState>,
    app: AppHandle,
) -> Result<LoadOptionsResult, String> {
    {
        let mut flag = state.cancel_flag.lock().unwrap();
        *flag = false;
    }

    // 检查文件是否已存在
    let mut config = load_data_source_list_config(&app)?;
    if config.data_sources.iter().any(|ds| ds.file_path == file_path) {
        return Err("该文件已经添加为数据源".to_string());
    }

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

    // 添加到数据源列表
    let id = uuid::Uuid::new_v4().to_string();
    config.data_sources.push(DataSourceConfig {
        id: id.clone(),
        file_path: result.file_path.clone(),
        file_name: file_name.clone(),
        loaded_at: chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string(),
        total_rows: result.total_rows,
    });
    
    // 设置为当前数据源
    config.current_id = Some(id.clone());
    
    // 保存配置
    save_data_source_list_config(&app, &config)?;

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

/// 设置数据源（兼容旧接口）
#[tauri::command]
async fn set_data_source(
    file_path: String,
    state: State<'_, AppState>,
    app: AppHandle,
) -> Result<LoadOptionsResult, String> {
    add_data_source(file_path, state, app).await
}

/// 获取数据源列表信息
#[tauri::command]
async fn get_data_source_list_info(
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<DataSourceListInfo, String> {
    let config = load_data_source_list_config(&app)?;
    
    // 检查缓存，更新total_rows
    let cache = state.data_cache.lock().unwrap();
    let cached_file_path = cache.as_ref().map(|c| &c.file_path);
    
    let data_sources: Vec<DataSourceInfo> = config.data_sources
        .into_iter()
        .map(|ds| {
            let total_rows = if cached_file_path == Some(&ds.file_path) {
                cache.as_ref().map(|c| c.cached_rows.len()).unwrap_or(ds.total_rows)
            } else {
                ds.total_rows
            };
            
            DataSourceInfo {
                id: ds.id,
                file_path: ds.file_path,
                file_name: ds.file_name,
                loaded_at: ds.loaded_at,
                total_rows,
            }
        })
        .collect();
    
    Ok(DataSourceListInfo {
        data_sources,
        current_id: config.current_id,
    })
}

/// 获取当前数据源信息（兼容旧接口）
#[tauri::command]
async fn get_data_source_info(
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<Option<DataSourceInfo>, String> {
    let list_info = get_data_source_list_info(app, state).await?;
    
    if let Some(current_id) = list_info.current_id {
        if let Some(current_ds) = list_info.data_sources.iter().find(|ds| ds.id == current_id) {
            return Ok(Some(current_ds.clone()));
        }
    }
    
    // 如果没有当前数据源，返回第一个
    if let Some(first_ds) = list_info.data_sources.first() {
        return Ok(Some(first_ds.clone()));
    }
    
    Ok(None)
}

/// 删除数据源
#[tauri::command]
async fn delete_data_source(
    data_source_id: String,
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<(), String> {
    let mut config = load_data_source_list_config(&app)?;
    
    // 检查是否是当前数据源
    let is_current = config.current_id.as_ref() == Some(&data_source_id);
    
    // 删除数据源
    config.data_sources.retain(|ds| ds.id != data_source_id);
    
    // 如果删除的是当前数据源，切换到第一个（如果有）
    if is_current {
        config.current_id = config.data_sources.first().map(|ds| ds.id.clone());
        
        // 如果还有数据源，加载新的当前数据源
        if let Some(new_current_id) = &config.current_id {
            // 加载新数据源
            let _ = load_data_source_by_id(new_current_id.clone(), state, app.clone()).await;
        } else {
            // 没有数据源了，清空缓存
            let mut cache = state.data_cache.lock().unwrap();
            *cache = None;
        }
    }
    
    save_data_source_list_config(&app, &config)?;
    Ok(())
}

/// 切换数据源
#[tauri::command]
async fn switch_data_source(
    data_source_id: String,
    app: AppHandle,
    state: State<'_, AppState>,
) -> Result<LoadOptionsResult, String> {
    load_data_source_by_id(data_source_id, state, app).await
}

/// 根据ID加载数据源
async fn load_data_source_by_id(
    data_source_id: String,
    state: State<'_, AppState>,
    app: AppHandle,
) -> Result<LoadOptionsResult, String> {
    let config = load_data_source_list_config(&app)?;
    
    let data_source = config.data_sources
        .iter()
        .find(|ds| ds.id == data_source_id)
        .ok_or("数据源不存在")?;
    
    // 检查缓存是否已经是这个数据源
    {
        let cache = state.data_cache.lock().unwrap();
        if let Some(ref data_cache) = *cache {
            if data_cache.file_path == data_source.file_path {
                // 已经是当前数据源，直接返回
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

                return Ok(LoadOptionsResult {
                    file_path: data_source.file_path.clone(),
                    file_name: data_source.file_name.clone(),
                    available_customers,
                    available_provinces,
                    available_cities,
                    available_districts,
                    available_regions,
                    total_rows: data_cache.cached_rows.len(),
                    load_time_ms: 0,
                });
            }
        }
    }
    
    // 需要加载数据源
    {
        let mut flag = state.cancel_flag.lock().unwrap();
        *flag = false;
    }

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
        let file_path = data_source.file_path.clone();
        move || {
            monthly_analysis::load_excel_file(&file_path, cancel_flag, progress_callback)
        }
    })
    .await
    .map_err(|e| format!("任务执行失败: {}", e))??;

    // 构建客户数据映射
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

    // 更新缓存
    {
        let mut cache = data_cache.lock().unwrap();
        *cache = Some(DataCache {
            file_path: result.file_path.clone(),
            cached_rows: result.cached_rows.clone(),
            customer_data_map,
        });
    }

    // 更新配置中的当前数据源
    let mut config = load_data_source_list_config(&app)?;
    config.current_id = Some(data_source_id);
    save_data_source_list_config(&app, &config)?;

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

/// 自动加载数据源（如果存在）
#[tauri::command]
async fn auto_load_data_source(
    state: State<'_, AppState>,
    app: AppHandle,
) -> Result<Option<LoadOptionsResult>, String> {
    let config = load_data_source_list_config(&app)?;
    
    let current_id = match config.current_id {
        Some(id) => id,
        None => {
            // 如果没有当前数据源，尝试使用第一个
            if let Some(first_ds) = config.data_sources.first() {
                first_ds.id.clone()
            } else {
                return Ok(None);
            }
        },
    };
    
    let current_ds = config.data_sources
        .iter()
        .find(|ds| ds.id == current_id)
        .ok_or("当前数据源不存在")?;

    // 如果缓存已存在且文件路径匹配，直接返回
    {
        let cache = state.data_cache.lock().unwrap();
        if let Some(ref data_cache) = *cache {
            if data_cache.file_path == current_ds.file_path {
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
                    file_path: current_ds.file_path.clone(),
                    file_name: current_ds.file_name.clone(),
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
    let result = load_data_source_by_id(current_id, state, app).await?;
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

/// 保存导出文件
#[tauri::command]
async fn save_export_file(
    file_path: String,
    content: String,
) -> Result<(), String> {
    fs::write(&file_path, content)
        .map_err(|e| format!("保存文件失败: {}", e))?;
    Ok(())
}

/// 获取订单明细（用于导出）
#[tauri::command]
async fn get_order_details(
    analysis_type: String,
    target: String,
    state: State<'_, AppState>,
) -> Result<Vec<monthly_analysis::CachedRow>, String> {
    let data_cache = state.data_cache.clone();
    
    let result = tokio::task::spawn_blocking(move || {
        let cache = data_cache.lock().unwrap();
        
        match cache.as_ref() {
            Some(data) => {
                let mut details: Vec<monthly_analysis::CachedRow> = Vec::new();
                
                for row in &data.cached_rows {
                    let matches = match analysis_type.as_str() {
                        "customer" => row.customer_code == target,
                        "province" => row.province.as_ref().map_or(false, |p| p == &target),
                        "city" => row.city.as_ref().map_or(false, |c| c == &target),
                        "district" => row.district.as_ref().map_or(false, |d| d == &target),
                        "region" => row.region.as_ref().map_or(false, |r| r == &target),
                        _ => false,
                    };
                    
                    if matches {
                        details.push(row.clone());
                    }
                }
                
                Ok(details)
            },
            None => Err("请先在首页导入数据源".to_string()),
        }
    })
    .await
    .map_err(|e| format!("任务执行失败: {}", e))??;

    Ok(result)
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
            add_data_source,
            get_data_source_info,
            get_data_source_list_info,
            delete_data_source,
            switch_data_source,
            auto_load_data_source,
            analyze_top20_cached,
            load_monthly_file,
            get_monthly_options,
            analyze_monthly_cached,
            clear_data_cache,
            cancel_analysis,
            save_export_file,
            get_order_details
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
