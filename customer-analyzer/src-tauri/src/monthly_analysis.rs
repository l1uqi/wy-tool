use calamine::{open_workbook, Reader, Xlsx, Xls, Data};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use std::path::Path;

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct CustomerOption {
    pub code: String,
    pub name: String,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct ProcessProgress {
    pub step: String,
    pub message: String,
    pub percent: u32,
    pub detail: String,
}

/// 缓存的行数据
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CachedRow {
    pub customer_code: String,
    pub customer_name: String,
    pub pay_amount: f64,
    pub recharge_deduction: f64,
    pub total_amount: f64,
    pub province: Option<String>,
    pub city: Option<String>,
    pub district: Option<String>,
    pub region: Option<String>,
    pub month: Option<String>,    // 格式 "2024-01"
}

/// 文件加载结果
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct FileLoadResult {
    pub file_path: String,
    pub cached_rows: Vec<CachedRow>,
    pub available_customers: Vec<CustomerOption>,
    pub available_provinces: Vec<String>,
    pub available_cities: Vec<String>,
    pub available_districts: Vec<String>,
    pub available_regions: Vec<String>,
    pub total_rows: usize,
    pub load_time_ms: u128,
}

/// 月度销售数据
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct MonthlySalesData {
    pub month: String,
    pub total_amount: f64,
    pub pay_amount: f64,
    pub recharge_deduction: f64,
    pub order_count: u32,
    pub mom_growth_rate: f64,  // 环比增长率
}

/// 月度分析结果
#[derive(Debug, Serialize, Deserialize)]
pub struct MonthlyAnalysisResult {
    pub analysis_type: String,
    pub target: String,
    pub target_name: String,
    pub monthly_data: Vec<MonthlySalesData>,
    pub total_amount: f64,
    pub total_orders: u32,
    pub process_time_ms: u128,
}

/// 加载Excel文件并缓存数据
pub fn load_excel_file<F>(
    file_path: &str,
    cancel_flag: Arc<Mutex<bool>>,
    progress_callback: F,
) -> Result<FileLoadResult, String>
where
    F: Fn(ProcessProgress) + Send + Sync,
{
    let start_time = std::time::Instant::now();
    let path = Path::new(file_path);

    let extension = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_default();

    progress_callback(ProcessProgress {
        step: "1/3".to_string(),
        message: "正在打开Excel文件...".to_string(),
        percent: 5,
        detail: format!("文件: {}", path.file_name().unwrap_or_default().to_string_lossy()),
    });

    if *cancel_flag.lock().unwrap() {
        return Err("用户取消操作".to_string());
    }

    // 读取Excel数据
    let rows: Vec<Vec<Data>> = match extension.as_str() {
        "xlsx" => {
            let mut workbook: Xlsx<_> = open_workbook(file_path)
                .map_err(|e| format!("无法打开Excel文件: {}", e))?;
            
            progress_callback(ProcessProgress {
                step: "1/3".to_string(),
                message: "正在解析Excel数据...".to_string(),
                percent: 15,
                detail: "读取工作表中...".to_string(),
            });

            let sheet_name = workbook.sheet_names().first()
                .ok_or("Excel文件没有工作表")?.clone();
            
            let range = workbook.worksheet_range(&sheet_name)
                .map_err(|e| format!("无法读取工作表: {}", e))?;
            
            range.rows().map(|r| r.to_vec()).collect()
        },
        "xls" => {
            let mut workbook: Xls<_> = open_workbook(file_path)
                .map_err(|e| format!("无法打开Excel文件: {}", e))?;
            
            progress_callback(ProcessProgress {
                step: "1/3".to_string(),
                message: "正在解析Excel数据...".to_string(),
                percent: 15,
                detail: "读取工作表中...".to_string(),
            });

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

    let total_rows = rows.len() - 1;
    
    progress_callback(ProcessProgress {
        step: "2/3".to_string(),
        message: "正在解析数据...".to_string(),
        percent: 40,
        detail: format!("共 {} 行数据", total_rows),
    });

    if *cancel_flag.lock().unwrap() {
        return Err("用户取消操作".to_string());
    }

    // 解析表头
    let header = &rows[0];
    let col_indices = find_column_indices(header)?;

    // 解析所有行并缓存
    let mut cached_rows: Vec<CachedRow> = Vec::with_capacity(total_rows);
    let mut customers_map: HashMap<String, String> = HashMap::new();
    let mut provinces_set: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut cities_set: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut districts_set: std::collections::HashSet<String> = std::collections::HashSet::new();
    let mut regions_set: std::collections::HashSet<String> = std::collections::HashSet::new();

    let data_rows: Vec<_> = rows.iter().skip(1).collect();
    let row_count = data_rows.len();
    
    for (i, row) in data_rows.iter().enumerate() {
        if let Some(parsed) = parse_row(row, &col_indices) {
            // 收集选项
            if !parsed.customer_code.is_empty() {
                customers_map.insert(parsed.customer_code.clone(), parsed.customer_name.clone());
            }
            if let Some(ref p) = parsed.province {
                if !p.is_empty() { provinces_set.insert(p.clone()); }
            }
            if let Some(ref c) = parsed.city {
                if !c.is_empty() { cities_set.insert(c.clone()); }
            }
            if let Some(ref d) = parsed.district {
                if !d.is_empty() { districts_set.insert(d.clone()); }
            }
            if let Some(ref r) = parsed.region {
                if !r.is_empty() { regions_set.insert(r.clone()); }
            }
            
            cached_rows.push(parsed);
        }
        
        if i % 10000 == 0 {
            let percent = 40 + ((i as f64 / row_count as f64) * 50.0) as u32;
            progress_callback(ProcessProgress {
                step: "2/3".to_string(),
                message: "正在解析数据...".to_string(),
                percent,
                detail: format!("已处理 {} / {} 行", i, row_count),
            });
        }
    }

    progress_callback(ProcessProgress {
        step: "3/3".to_string(),
        message: "正在整理数据...".to_string(),
        percent: 95,
        detail: "生成选项列表...".to_string(),
    });

    // 整理选项
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

    let load_time_ms = start_time.elapsed().as_millis();

    progress_callback(ProcessProgress {
        step: "完成".to_string(),
        message: "数据加载完成！".to_string(),
        percent: 100,
        detail: format!("耗时 {}ms，缓存 {} 行数据", load_time_ms, cached_rows.len()),
    });

    Ok(FileLoadResult {
        file_path: file_path.to_string(),
        cached_rows,
        available_customers,
        available_provinces,
        available_cities,
        available_districts,
        available_regions,
        total_rows,
        load_time_ms,
    })
}

/// 基于缓存数据进行月度分析
pub fn analyze_from_cache(
    cached_rows: &[CachedRow],
    analysis_type: &str,
    target: &str,
) -> Result<MonthlyAnalysisResult, String> {
    let start_time = std::time::Instant::now();
    
    if target.is_empty() {
        return Err("请选择分析目标".to_string());
    }

    // 按月份汇总
    let mut monthly_map: HashMap<String, MonthlySalesData> = HashMap::new();
    let mut target_name = String::new();
    
    for row in cached_rows {
        let matches = match analysis_type {
            "customer" => {
                if row.customer_code == target {
                    if target_name.is_empty() && !row.customer_name.is_empty() {
                        target_name = row.customer_name.clone();
                    }
                    true
                } else {
                    false
                }
            },
            "province" => row.province.as_ref().map_or(false, |p| p == target),
            "city" => row.city.as_ref().map_or(false, |c| c == target),
            "district" => row.district.as_ref().map_or(false, |d| d == target),
            "region" => row.region.as_ref().map_or(false, |r| r == target),
            _ => false,
        };

        if matches {
            let month = row.month.clone().unwrap_or_else(|| "未知月份".to_string());
            
            monthly_map
                .entry(month.clone())
                .and_modify(|data| {
                    data.total_amount += row.total_amount;
                    data.pay_amount += row.pay_amount;
                    data.recharge_deduction += row.recharge_deduction;
                    data.order_count += 1;
                })
                .or_insert(MonthlySalesData {
                    month,
                    total_amount: row.total_amount,
                    pay_amount: row.pay_amount,
                    recharge_deduction: row.recharge_deduction,
                    order_count: 1,
                    mom_growth_rate: 0.0,
                });
        }
    }

    // 按月份排序
    let mut monthly_data: Vec<MonthlySalesData> = monthly_map.into_values().collect();
    monthly_data.sort_by(|a, b| a.month.cmp(&b.month));

    // 计算环比增长率
    for i in 0..monthly_data.len() {
        if i > 0 {
            let prev_amount = monthly_data[i - 1].total_amount;
            let curr_amount = monthly_data[i].total_amount;
            if prev_amount > 0.0 {
                monthly_data[i].mom_growth_rate = ((curr_amount - prev_amount) / prev_amount) * 100.0;
            } else if curr_amount > 0.0 {
                monthly_data[i].mom_growth_rate = 100.0;
            }
        }
    }

    let total_amount: f64 = monthly_data.iter().map(|d| d.total_amount).sum();
    let total_orders: u32 = monthly_data.iter().map(|d| d.order_count).sum();

    if target_name.is_empty() {
        target_name = target.to_string();
    }

    let process_time_ms = start_time.elapsed().as_millis();

    Ok(MonthlyAnalysisResult {
        analysis_type: analysis_type.to_string(),
        target: target.to_string(),
        target_name,
        monthly_data,
        total_amount,
        total_orders,
        process_time_ms,
    })
}

#[derive(Debug)]
struct ColumnIndices {
    customer_code: usize,
    customer_name: Option<usize>,
    pay_amount: usize,
    recharge_deduction: usize,
    province: Option<usize>,
    city: Option<usize>,
    district: Option<usize>,
    region: Option<usize>,
    date: Option<usize>,
}

fn find_column_indices(header: &[Data]) -> Result<ColumnIndices, String> {
    let mut customer_code_idx = None;
    let mut customer_name_idx = None;
    let mut pay_amount_idx = None;
    let mut recharge_deduction_idx = None;
    let mut province_idx = None;
    let mut city_idx = None;
    let mut district_idx = None;
    let mut region_idx = None;
    let mut date_idx = None;

    for (idx, cell) in header.iter().enumerate() {
        let col_name = data_to_string(cell).trim().to_string();
        match col_name.as_str() {
            "客户编码" => customer_code_idx = Some(idx),
            "客户名称" | "客户" => customer_name_idx = Some(idx),
            "支付金额" => pay_amount_idx = Some(idx),
            "充值抵扣" => recharge_deduction_idx = Some(idx),
            "省" | "省份" => province_idx = Some(idx),
            "市" | "城市" => city_idx = Some(idx),
            "区" | "区县" | "县" => district_idx = Some(idx),
            "地区" | "区域" => region_idx = Some(idx),
            "日期" | "订单日期" | "下单日期" | "创建时间" | "下单时间" | "支付时间" | "付款时间" | "交易时间" | "时间" | "成交时间" | "签约时间" | 
            "出库时间" | "出库日期" | "发货时间" | "发货日期" | "完成时间" | "完成日期" | "结算时间" | "结算日期" => date_idx = Some(idx),
            _ => {}
        }
    }

    let customer_code = customer_code_idx.ok_or("缺少必需列: 客户编码")?;
    let pay_amount = pay_amount_idx.ok_or("缺少必需列: 支付金额")?;
    let recharge_deduction = recharge_deduction_idx.ok_or("缺少必需列: 充值抵扣")?;

    Ok(ColumnIndices {
        customer_code,
        customer_name: customer_name_idx,
        pay_amount,
        recharge_deduction,
        province: province_idx,
        city: city_idx,
        district: district_idx,
        region: region_idx,
        date: date_idx,
    })
}

fn parse_row(row: &[Data], indices: &ColumnIndices) -> Option<CachedRow> {
    let customer_code = row
        .get(indices.customer_code)
        .map(|v| data_to_string(v).trim().to_string())?;
    
    if customer_code.is_empty() {
        return None;
    }

    let customer_name = indices
        .customer_name
        .and_then(|idx| row.get(idx))
        .map(|v| data_to_string(v).trim().to_string())
        .unwrap_or_default();

    let pay_amount = row
        .get(indices.pay_amount)
        .map(|v| parse_number(v))
        .unwrap_or(0.0);

    let recharge_deduction = row
        .get(indices.recharge_deduction)
        .map(|v| parse_number(v))
        .unwrap_or(0.0);

    let total_amount = pay_amount + recharge_deduction;

    let province = indices
        .province
        .and_then(|idx| row.get(idx))
        .map(|v| data_to_string(v).trim().to_string())
        .filter(|s| !s.is_empty());

    let city = indices
        .city
        .and_then(|idx| row.get(idx))
        .map(|v| data_to_string(v).trim().to_string())
        .filter(|s| !s.is_empty());

    let district = indices
        .district
        .and_then(|idx| row.get(idx))
        .map(|v| data_to_string(v).trim().to_string())
        .filter(|s| !s.is_empty());

    // 组合地区
    let region = if province.is_some() || city.is_some() || district.is_some() {
        let parts: Vec<&str> = [
            province.as_deref(),
            city.as_deref(), 
            district.as_deref()
        ].into_iter().flatten().collect();
        
        if !parts.is_empty() {
            Some(parts.join("-"))
        } else {
            None
        }
    } else {
        indices
            .region
            .and_then(|idx| row.get(idx))
            .map(|v| data_to_string(v).trim().to_string())
            .filter(|s| !s.is_empty())
    };

    // 解析月份
    let month = indices
        .date
        .and_then(|idx| row.get(idx))
        .and_then(|v| extract_month(v));

    Some(CachedRow {
        customer_code,
        customer_name,
        pay_amount,
        recharge_deduction,
        total_amount,
        province,
        city,
        district,
        region,
        month,
    })
}

/// 提取月份
fn extract_month(value: &Data) -> Option<String> {
    match value {
        Data::DateTime(dt) => {
            let days = dt.as_f64();
            if days > 0.0 {
                let base_date = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)?;
                let date = base_date + chrono::Duration::days(days as i64);
                Some(date.format("%Y-%m").to_string())
            } else {
                None
            }
        },
        Data::DateTimeIso(s) => {
            if s.len() >= 7 {
                Some(s[0..7].to_string())
            } else {
                None
            }
        },
        Data::String(s) => parse_month_from_string(s),
        Data::Float(f) => {
            let days = *f;
            if days > 0.0 && days < 100000.0 {
                let base_date = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)?;
                let date = base_date + chrono::Duration::days(days as i64);
                Some(date.format("%Y-%m").to_string())
            } else {
                None
            }
        },
        _ => None,
    }
}

fn parse_month_from_string(s: &str) -> Option<String> {
    let s = s.trim();
    
    // 2024-01-15, 2024/01/15, 2024.01.15
    if let Some(cap) = regex_lite::Regex::new(r"^(\d{4})[-/.](\d{1,2})")
        .ok()
        .and_then(|re| re.captures(s))
    {
        let year = cap.get(1)?.as_str();
        let month = cap.get(2)?.as_str();
        return Some(format!("{}-{:0>2}", year, month));
    }
    
    // 2024-01-15 10:30:00 格式（带时间）
    if let Some(cap) = regex_lite::Regex::new(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})")
        .ok()
        .and_then(|re| re.captures(s))
    {
        let year = cap.get(1)?.as_str();
        let month = cap.get(2)?.as_str();
        return Some(format!("{}-{:0>2}", year, month));
    }
    
    // 2024年1月15日 或 2024年1月 格式
    if let Some(cap) = regex_lite::Regex::new(r"^(\d{4})年(\d{1,2})月")
        .ok()
        .and_then(|re| re.captures(s))
    {
        let year = cap.get(1)?.as_str();
        let month = cap.get(2)?.as_str();
        return Some(format!("{}-{:0>2}", year, month));
    }
    
    // 20240115 格式
    if s.len() >= 6 && s.chars().take(6).all(|c| c.is_ascii_digit()) {
        let year = &s[0..4];
        let month = &s[4..6];
        return Some(format!("{}-{}", year, month));
    }
    
    // 尝试使用 chrono 解析常见日期格式
    if let Ok(date) = chrono::NaiveDate::parse_from_str(s, "%Y-%m-%d") {
        return Some(date.format("%Y-%m").to_string());
    }
    if let Ok(date) = chrono::NaiveDate::parse_from_str(s, "%Y/%m/%d") {
        return Some(date.format("%Y-%m").to_string());
    }
    if let Ok(date) = chrono::NaiveDate::parse_from_str(s, "%Y.%m.%d") {
        return Some(date.format("%Y-%m").to_string());
    }
    
    None
}

fn data_to_string(value: &Data) -> String {
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

fn parse_number(value: &Data) -> f64 {
    match value {
        Data::Float(f) => *f,
        Data::Int(i) => *i as f64,
        Data::String(s) => s.trim().parse().unwrap_or(0.0),
        _ => 0.0,
    }
}
