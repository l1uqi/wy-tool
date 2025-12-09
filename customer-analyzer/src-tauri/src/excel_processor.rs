use calamine::{open_workbook, Reader, Xlsx, Xls, Data};
use rayon::prelude::*;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use std::path::Path;

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct CustomerData {
    pub customer_code: String,
    pub customer_name: String,
    pub pay_amount: f64,
    pub recharge_deduction: f64,
    pub total_amount: f64,
    pub order_count: u32,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct AnalysisResult {
    pub top20: Vec<CustomerData>,
    pub total_customers: usize,
    pub total_amount: f64,
    pub top20_amount: f64,
    pub total_rows: usize,
    pub process_time_ms: u128,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct ProcessProgress {
    pub step: String,
    pub message: String,
    pub percent: u32,
    pub detail: String,
}

pub fn process_excel_file<F>(
    file_path: &str,
    cancel_flag: Arc<Mutex<bool>>,
    progress_callback: F,
) -> Result<AnalysisResult, String>
where
    F: Fn(ProcessProgress) + Send + Sync,
{
    let start_time = std::time::Instant::now();
    let path = Path::new(file_path);

    // 检查文件扩展名
    let extension = path
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_lowercase())
        .unwrap_or_default();

    progress_callback(ProcessProgress {
        step: "1/4".to_string(),
        message: "正在打开Excel文件...".to_string(),
        percent: 5,
        detail: format!("文件: {}", path.file_name().unwrap_or_default().to_string_lossy()),
    });

    // 检查是否取消
    if *cancel_flag.lock().unwrap() {
        return Err("用户取消操作".to_string());
    }

    // 根据扩展名打开不同类型的Excel文件
    let rows: Vec<Vec<Data>> = match extension.as_str() {
        "xlsx" => {
            let mut workbook: Xlsx<_> = open_workbook(file_path)
                .map_err(|e| format!("无法打开Excel文件: {}", e))?;
            
            progress_callback(ProcessProgress {
                step: "2/4".to_string(),
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
                step: "2/4".to_string(),
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

    // 检查是否取消
    if *cancel_flag.lock().unwrap() {
        return Err("用户取消操作".to_string());
    }

    let total_rows = rows.len() - 1; // 减去表头
    
    progress_callback(ProcessProgress {
        step: "2/4".to_string(),
        message: "数据读取完成".to_string(),
        percent: 30,
        detail: format!("共 {} 行数据", total_rows),
    });

    // 解析表头，找到列索引
    let header = &rows[0];
    let col_indices = find_column_indices(header)?;

    progress_callback(ProcessProgress {
        step: "3/4".to_string(),
        message: "正在分析客户数据...".to_string(),
        percent: 40,
        detail: "使用多线程并行处理...".to_string(),
    });

    // 检查是否取消
    if *cancel_flag.lock().unwrap() {
        return Err("用户取消操作".to_string());
    }

    // 使用Rayon并行处理数据
    let data_rows: Vec<_> = rows.iter().skip(1).collect();
    let chunk_size = (data_rows.len() / rayon::current_num_threads().max(1)).max(1000);
    
    // 分块并行处理
    let partial_maps: Vec<HashMap<String, CustomerData>> = data_rows
        .par_chunks(chunk_size)
        .map(|chunk| {
            let mut local_map: HashMap<String, CustomerData> = HashMap::new();
            
            for row in chunk {
                if let Some(customer) = parse_row(row, &col_indices) {
                    local_map
                        .entry(customer.customer_code.clone())
                        .and_modify(|existing| {
                            existing.pay_amount += customer.pay_amount;
                            existing.recharge_deduction += customer.recharge_deduction;
                            existing.total_amount += customer.total_amount;
                            existing.order_count += 1;
                            if existing.customer_name.is_empty() && !customer.customer_name.is_empty() {
                                existing.customer_name = customer.customer_name.clone();
                            }
                        })
                        .or_insert(customer);
                }
            }
            
            local_map
        })
        .collect();

    // 检查是否取消
    if *cancel_flag.lock().unwrap() {
        return Err("用户取消操作".to_string());
    }

    progress_callback(ProcessProgress {
        step: "3/4".to_string(),
        message: "正在合并处理结果...".to_string(),
        percent: 70,
        detail: format!("处理了 {} 个数据块", partial_maps.len()),
    });

    // 合并所有部分结果
    let mut customer_map: HashMap<String, CustomerData> = HashMap::new();
    for partial in partial_maps {
        for (code, customer) in partial {
            customer_map
                .entry(code)
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
    }

    progress_callback(ProcessProgress {
        step: "4/4".to_string(),
        message: "正在生成排行榜...".to_string(),
        percent: 85,
        detail: format!("发现 {} 个不同客户", customer_map.len()),
    });

    // 转换为Vec并排序
    let mut customers: Vec<CustomerData> = customer_map.into_values().collect();
    customers.par_sort_by(|a, b| {
        b.total_amount
            .partial_cmp(&a.total_amount)
            .unwrap_or(std::cmp::Ordering::Equal)
    });

    // 计算统计数据
    let total_amount: f64 = customers.iter().map(|c| c.total_amount).sum();
    let top20: Vec<CustomerData> = customers.iter().take(20).cloned().collect();
    let top20_amount: f64 = top20.iter().map(|c| c.total_amount).sum();

    let process_time_ms = start_time.elapsed().as_millis();

    progress_callback(ProcessProgress {
        step: "完成".to_string(),
        message: "分析完成！".to_string(),
        percent: 100,
        detail: format!("耗时 {}ms", process_time_ms),
    });

    Ok(AnalysisResult {
        top20,
        total_customers: customers.len(),
        total_amount,
        top20_amount,
        total_rows,
        process_time_ms,
    })
}

#[derive(Debug)]
struct ColumnIndices {
    customer_code: usize,
    customer_name: Option<usize>,
    pay_amount: usize,
    recharge_deduction: usize,
}

fn find_column_indices(header: &[Data]) -> Result<ColumnIndices, String> {
    let mut customer_code_idx = None;
    let mut customer_name_idx = None;
    let mut pay_amount_idx = None;
    let mut recharge_deduction_idx = None;

    for (idx, cell) in header.iter().enumerate() {
        let col_name = data_to_string(cell).trim().to_string();
        match col_name.as_str() {
            "客户编码" => customer_code_idx = Some(idx),
            "客户名称" | "客户" => customer_name_idx = Some(idx),
            "支付金额" => pay_amount_idx = Some(idx),
            "充值抵扣" => recharge_deduction_idx = Some(idx),
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
    })
}

fn parse_row(row: &[Data], indices: &ColumnIndices) -> Option<CustomerData> {
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

    Some(CustomerData {
        customer_code,
        customer_name,
        pay_amount,
        recharge_deduction,
        total_amount,
        order_count: 1,
    })
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
