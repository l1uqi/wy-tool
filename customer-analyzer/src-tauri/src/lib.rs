mod excel_processor;
mod monthly_analysis;
mod out_of_policy;

use excel_processor::{AnalysisResult, ProcessProgress, CustomerData};
use monthly_analysis::{MonthlyAnalysisResult, CachedRow, CustomerOption};
use out_of_policy::{OutOfPolicyResult};
use std::sync::{Arc, Mutex};
use tauri::{AppHandle, Emitter, State};
use serde::{Deserialize, Serialize};
use std::path::PathBuf;
use std::fs;
use std::env;
use calamine::Data;

// 全局状态
struct AppState {
    cancel_flag: Arc<Mutex<bool>>,
    data_cache: Arc<Mutex<Option<DataCache>>>,
}

#[derive(Clone, Serialize, Deserialize)]
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

/// 获取应用数据目录
fn get_app_data_dir() -> PathBuf {
    dirs::data_dir()
        .map(|d| d.join("customer-analyzer"))
        .unwrap_or_else(|| {
            env::current_dir().unwrap_or_else(|_| PathBuf::from("."))
        })
}

/// 获取配置文件路径
fn get_config_path(_app: &AppHandle) -> PathBuf {
    let app_data_dir = get_app_data_dir();
    std::fs::create_dir_all(&app_data_dir).unwrap_or_default();
    app_data_dir.join("data_source.json")
}

/// 获取数据源缓存文件路径
fn get_cache_path(data_source_id: &str) -> PathBuf {
    let app_data_dir = get_app_data_dir();
    std::fs::create_dir_all(&app_data_dir).unwrap_or_default();
    app_data_dir.join(format!("cache_{}.json", data_source_id))
}

/// 保存数据源缓存到文件
fn save_data_cache(data_source_id: &str, cache: &DataCache) -> Result<(), String> {
    let cache_path = get_cache_path(data_source_id);
    let json = serde_json::to_string_pretty(cache)
        .map_err(|e| format!("序列化缓存失败: {}", e))?;
    fs::write(&cache_path, json)
        .map_err(|e| format!("保存缓存文件失败: {}", e))?;
    Ok(())
}

/// 从文件加载数据源缓存
fn load_data_cache(data_source_id: &str) -> Result<Option<DataCache>, String> {
    let cache_path = get_cache_path(data_source_id);
    
    if !cache_path.exists() {
        return Ok(None);
    }
    
    let content = fs::read_to_string(&cache_path)
        .map_err(|e| format!("读取缓存文件失败: {}", e))?;
    
    let cache: DataCache = serde_json::from_str(&content)
        .map_err(|e| format!("解析缓存文件失败: {}", e))?;
    
    Ok(Some(cache))
}

/// 删除数据源缓存文件
fn delete_data_cache(data_source_id: &str) -> Result<(), String> {
    let cache_path = get_cache_path(data_source_id);
    if cache_path.exists() {
        fs::remove_file(&cache_path)
            .map_err(|e| format!("删除缓存文件失败: {}", e))?;
    }
    Ok(())
}

/// 合并多个数据源的缓存
fn merge_data_caches(data_source_ids: Vec<String>, _app: &AppHandle) -> Result<DataCache, String> {
    if data_source_ids.is_empty() {
        return Err("至少需要选择一个数据源".to_string());
    }
    
    let mut merged_rows: Vec<CachedRow> = Vec::new();
    let mut merged_customer_map: std::collections::HashMap<String, CustomerData> = 
        std::collections::HashMap::new();
    let mut merged_file_paths: Vec<String> = Vec::new();
    
    for id in &data_source_ids {
        match load_data_cache(id)? {
            Some(cache) => {
                merged_file_paths.push(cache.file_path.clone());
                merged_rows.extend(cache.cached_rows.clone());
                
                // 合并客户数据映射
                for (code, customer) in cache.customer_data_map {
                    merged_customer_map
                        .entry(code.clone())
                        .and_modify(|existing| {
                            existing.pay_amount += customer.pay_amount;
                            existing.recharge_deduction += customer.recharge_deduction;
                            existing.total_amount += customer.total_amount;
                            existing.order_count += customer.order_count;
                            if existing.customer_name.is_empty() && !customer.customer_name.is_empty() {
                                existing.customer_name = customer.customer_name.clone();
                            }
                        })
                        .or_insert(customer);
                }
            },
            None => {
                return Err(format!("数据源 {} 的缓存不存在，请先加载该数据源", id));
            }
        }
    }
    
    Ok(DataCache {
        file_path: merged_file_paths.join("; "),
        cached_rows: merged_rows,
        customer_data_map: merged_customer_map,
    })
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
    // 预分配容量以提高性能
    let estimated_customers = (result.cached_rows.len() / 10).max(100).min(10000);
    let mut customer_data_map: std::collections::HashMap<String, CustomerData> = 
        std::collections::HashMap::with_capacity(estimated_customers);
    
    // 优化：减少克隆操作，直接使用引用
    for row in &result.cached_rows {
        customer_data_map
            .entry(row.customer_code.clone())
            .and_modify(|existing| {
                existing.pay_amount += row.pay_amount;
                existing.recharge_deduction += row.recharge_deduction;
                existing.total_amount += row.total_amount;
                existing.order_count += 1;
                // 只在需要时更新客户名称
                if existing.customer_name.is_empty() && !row.customer_name.is_empty() {
                    existing.customer_name.clone_from(&row.customer_name);
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

    // 添加到数据源列表
    let id = uuid::Uuid::new_v4().to_string();
    let file_name = std::path::Path::new(&result.file_path)
        .file_name()
        .and_then(|n| n.to_str())
        .unwrap_or("未知文件")
        .to_string();
    
    config.data_sources.push(DataSourceConfig {
        id: id.clone(),
        file_path: result.file_path.clone(),
        file_name: file_name.clone(),
        loaded_at: chrono::Local::now().format("%Y-%m-%d %H:%M:%S").to_string(),
        total_rows: result.total_rows,
    });
    
    // 设置为当前数据源
    config.current_id = Some(id.clone());
    
    // 创建缓存对象
    let cache_obj = DataCache {
        file_path: result.file_path.clone(),
        cached_rows: result.cached_rows.clone(),
        customer_data_map,
    };
    
    // 保存缓存到文件（持久化）
    save_data_cache(&id, &cache_obj)?;
    
    // 缓存到内存
    {
        let mut cache = data_cache.lock().unwrap();
        *cache = Some(cache_obj);
    }
    
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
    
    // 删除数据源缓存文件
    let _ = delete_data_cache(&data_source_id);
    
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
    
    // 检查内存缓存是否已经是这个数据源
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
    
    // 尝试从文件加载缓存（持久化）
    if let Ok(Some(cached_data)) = load_data_cache(&data_source_id) {
        // 验证文件路径是否匹配（防止文件被移动或重命名）
        if cached_data.file_path == data_source.file_path {
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

            for row in &cached_data.cached_rows {
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

            // 保存总行数（在移动之前）
            let total_rows = cached_data.cached_rows.len();

            // 加载到内存缓存
            {
                let mut cache = state.data_cache.lock().unwrap();
                *cache = Some(cached_data);
            }

            return Ok(LoadOptionsResult {
                file_path: data_source.file_path.clone(),
                file_name: data_source.file_name.clone(),
                available_customers,
                available_provinces,
                available_cities,
                available_districts,
                available_regions,
                total_rows,
                load_time_ms: 0,
            });
        }
    }
    
    // 缓存不存在或文件路径不匹配，需要从Excel文件重新加载
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

    // 创建缓存对象
    let cache_obj = DataCache {
        file_path: result.file_path.clone(),
        cached_rows: result.cached_rows.clone(),
        customer_data_map,
    };
    
    // 保存缓存到文件（持久化）
    save_data_cache(&data_source_id, &cache_obj)?;
    
    // 更新内存缓存
    {
        let mut cache = data_cache.lock().unwrap();
        *cache = Some(cache_obj);
    }

    // 更新配置中的当前数据源
    let mut config = load_data_source_list_config(&app)?;
    config.current_id = Some(data_source_id.clone());
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

/// 前20大客户分析（支持多数据源合并）
#[tauri::command]
async fn analyze_top20_multi(
    data_source_ids: Vec<String>,
    _state: State<'_, AppState>,
    app: AppHandle,
) -> Result<AnalysisResult, String> {
    // 合并多个数据源的缓存
    let merged_cache = merge_data_caches(data_source_ids, &app)?;
    
    let result = tokio::task::spawn_blocking(move || -> Result<AnalysisResult, String> {
        let mut customers: Vec<CustomerData> = 
            merged_cache.customer_data_map.values().cloned().collect();
        
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
            total_rows: merged_cache.cached_rows.len(),
            process_time_ms: 0,
        })
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

/// 基于合并的多个数据源执行月度分析
#[tauri::command]
async fn analyze_monthly_multi(
    data_source_ids: Vec<String>,
    analysis_type: String,
    target: String,
    app: AppHandle,
) -> Result<MonthlyAnalysisResult, String> {
    // 合并多个数据源的缓存
    let merged_cache = merge_data_caches(data_source_ids, &app)?;
    
    let result = tokio::task::spawn_blocking(move || {
        monthly_analysis::analyze_from_cache(
            &merged_cache.cached_rows, 
            &analysis_type, 
            &target
        )
    })
    .await
    .map_err(|e| format!("任务执行失败: {}", e))??;

    Ok(result)
}

/// 获取合并后的月度分析选项
#[tauri::command]
async fn get_monthly_options_multi(
    data_source_ids: Vec<String>,
    app: AppHandle,
) -> Result<LoadOptionsResult, String> {
    // 合并多个数据源的缓存
    let merged_cache = merge_data_caches(data_source_ids, &app)?;
    
    // 构建选项
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

    for row in &merged_cache.cached_rows {
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

    Ok(LoadOptionsResult {
        file_path: merged_cache.file_path,
        file_name: "合并数据源".to_string(),
        available_customers,
        available_provinces,
        available_cities,
        available_districts,
        available_regions,
        total_rows: merged_cache.cached_rows.len(),
        load_time_ms: 0,
    })
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

/// 保存 Excel 文件（base64 编码）
#[tauri::command]
async fn save_excel_file(
    file_path: String,
    content: String,
) -> Result<(), String> {
    use base64::{Engine as _, engine::general_purpose};
    
    // 解码 base64 字符串为二进制数据
    let bytes = general_purpose::STANDARD
        .decode(content)
        .map_err(|e| format!("Base64 解码失败: {}", e))?;
    
    // 保存二进制数据到文件
    fs::write(&file_path, bytes)
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

/// 客户采购额月度数据
#[derive(Debug, Serialize, Deserialize)]
struct CustomerMonthlyPurchase {
    month: String,
    total_amount: f64,
}

/// 客户采购额数据
#[derive(Debug, Serialize, Deserialize)]
struct CustomerPurchaseData {
    customer_code: String,
    customer_name: String,
    monthly_data: Vec<CustomerMonthlyPurchase>,
    total_amount: f64,
}

/// 客户采购额计算结果
#[derive(Debug, Serialize, Deserialize)]
struct CustomerPurchaseResult {
    customer_data: Vec<CustomerPurchaseData>,
    total_customers: usize,
    total_amount: f64,
}

/// 加载客户编码Excel文件（返回完整数据）
#[tauri::command]
async fn load_customer_codes(
    file_path: String,
) -> Result<serde_json::Value, String> {
    use calamine::{open_workbook, Reader, Xlsx, Xls, Data};
    use std::path::Path;
    
    let path = Path::new(&file_path);
    let extension = path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();
    
    let mut customer_codes: Vec<String> = Vec::new();
    
    let rows: Vec<Vec<Data>> = match extension.as_str() {
        "xlsx" => {
            let mut workbook: Xlsx<_> = open_workbook(&file_path)
                .map_err(|e| format!("无法打开Excel文件: {}", e))?;
            
            let sheet_name = workbook.sheet_names().first()
                .ok_or("Excel文件没有工作表")?.clone();
            
            let range = workbook.worksheet_range(&sheet_name)
                .map_err(|e| format!("无法读取工作表: {}", e))?;
            
            range.rows().map(|r| r.to_vec()).collect()
        },
        "xls" => {
            let mut workbook: Xls<_> = open_workbook(&file_path)
                .map_err(|e| format!("无法打开Excel文件: {}", e))?;
            
            let sheet_name = workbook.sheet_names().first()
                .ok_or("Excel文件没有工作表")?.clone();
            
            let range = workbook.worksheet_range(&sheet_name)
                .map_err(|e| format!("无法读取工作表: {}", e))?;
            
            range.rows().map(|r| r.to_vec()).collect()
        },
        _ => return Err(format!("不支持的文件格式: {}", extension)),
    };
    
    if rows.is_empty() {
        return Err("Excel文件为空".to_string());
    }
    
    // 查找客户编码列索引
    let header = &rows[0];
    let mut customer_code_idx = None;
    
    for (idx, cell) in header.iter().enumerate() {
        let col_name = purchase_data_to_string(cell).trim().to_string();
        if col_name == "客户编码" || col_name == "客户代码" || col_name == "编码" {
            customer_code_idx = Some(idx);
            break;
        }
    }
    
    let customer_code_idx = customer_code_idx.ok_or("未找到'客户编码'列")?;
    
    // 提取客户编码和完整行数据
    let mut excel_rows: Vec<Vec<String>> = Vec::new();
    for row in rows.iter().skip(1) {
        if let Some(cell) = row.get(customer_code_idx) {
            let code = purchase_data_to_string(cell).trim().to_string();
            if !code.is_empty() {
                customer_codes.push(code.clone());
                // 保存完整的行数据（转换为字符串）
                let row_data: Vec<String> = row.iter().map(|d| purchase_data_to_string(d)).collect();
                excel_rows.push(row_data);
            }
        }
    }
    
    // 转换表头
    let headers: Vec<String> = header.iter().map(|d| purchase_data_to_string(d)).collect();
    
    Ok(serde_json::json!({
        "customer_codes": customer_codes,
        "headers": headers,
        "rows": excel_rows,
        "customer_code_index": customer_code_idx
    }))
}

fn purchase_data_to_string(value: &calamine::Data) -> String {
    match value {
        Data::Int(i) => i.to_string(),
        Data::Float(f) => f.to_string(),
        Data::String(s) => s.clone(),
        Data::Bool(b) => b.to_string(),
        Data::DateTime(dt) => dt.to_string(),
        Data::DateTimeIso(s) => s.clone(),
        Data::DurationIso(s) => s.clone(),
        Data::Error(e) => format!("{:?}", e),
        Data::Empty => String::new(),
    }
}

/// 加载政策外开单Excel文件
#[tauri::command]
async fn load_out_of_policy_excel(
    file_path: String,
) -> Result<OutOfPolicyResult, String> {
    tokio::task::spawn_blocking(move || {
        out_of_policy::load_out_of_policy_file(&file_path)
    })
    .await
    .map_err(|e| format!("任务执行失败: {}", e))?
}

/// 计算客户采购额
#[tauri::command]
async fn calculate_customer_purchase(
    data_source_ids: Vec<String>,
    customer_codes: Vec<String>,
    app: AppHandle,
) -> Result<CustomerPurchaseResult, String> {
    // 合并多个数据源的缓存
    let merged_cache = merge_data_caches(data_source_ids, &app)?;
    
    let result = tokio::task::spawn_blocking(move || -> Result<CustomerPurchaseResult, String> {
        // 创建客户编码集合用于快速查找
        let customer_code_set: std::collections::HashSet<String> = 
            customer_codes.iter().cloned().collect();
        
        // 按客户编码和月份分组统计
        let mut customer_map: std::collections::HashMap<String, CustomerPurchaseData> = 
            std::collections::HashMap::new();
        
        for row in &merged_cache.cached_rows {
            // 只处理在客户编码列表中的客户
            if !customer_code_set.contains(&row.customer_code) {
                continue;
            }
            
            let month = row.month.clone().unwrap_or_else(|| "未知月份".to_string());
            
            customer_map
                .entry(row.customer_code.clone())
                .and_modify(|data| {
                    // 查找或创建该月份的记录
                    let monthly_entry = data.monthly_data
                        .iter_mut()
                        .find(|m| m.month == month);
                    
                    if let Some(entry) = monthly_entry {
                        entry.total_amount += row.total_amount;
                    } else {
                        data.monthly_data.push(CustomerMonthlyPurchase {
                            month: month.clone(),
                            total_amount: row.total_amount,
                        });
                    }
                    
                    data.total_amount += row.total_amount;
                    
                    // 更新客户名称（如果为空）
                    if data.customer_name.is_empty() && !row.customer_name.is_empty() {
                        data.customer_name = row.customer_name.clone();
                    }
                })
                .or_insert_with(|| {
                    let mut monthly_data = Vec::new();
                    monthly_data.push(CustomerMonthlyPurchase {
                        month: month.clone(),
                        total_amount: row.total_amount,
                    });
                    
                    CustomerPurchaseData {
                        customer_code: row.customer_code.clone(),
                        customer_name: row.customer_name.clone(),
                        monthly_data,
                        total_amount: row.total_amount,
                    }
                });
        }
        
        // 转换为Vec并排序
        let mut customer_data: Vec<CustomerPurchaseData> = customer_map.into_values().collect();
        customer_data.sort_by(|a, b| a.customer_code.cmp(&b.customer_code));
        
        // 对每个客户的月度数据按月份排序
        for customer in &mut customer_data {
            customer.monthly_data.sort_by(|a, b| a.month.cmp(&b.month));
        }
        
        let total_customers = customer_data.len();
        let total_amount: f64 = customer_data.iter().map(|c| c.total_amount).sum();
        
        Ok(CustomerPurchaseResult {
            customer_data,
            total_customers,
            total_amount,
        })
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
            analyze_top20_multi,
            load_monthly_file,
            get_monthly_options,
            get_monthly_options_multi,
            analyze_monthly_cached,
            analyze_monthly_multi,
            clear_data_cache,
            cancel_analysis,
            save_export_file,
            save_excel_file,
            get_order_details,
            load_customer_codes,
            calculate_customer_purchase,
            load_out_of_policy_excel
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
