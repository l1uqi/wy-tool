use calamine::{open_workbook, Reader, Xlsx, Data};
use serde::{Deserialize, Serialize};
use std::path::Path;
use rayon::prelude::*;
use chrono::{Datelike, NaiveDate};
use std::collections::HashMap;
use rust_xlsxwriter::Workbook;

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct OutOfPolicyRow {
    pub order_date: String,
    pub sales_order_no: String,
    pub customer_code: String,
    pub customer_name: String,
    pub product_code: String,
    pub generic_name: String,
    pub sales_price: f64,
    pub settlement_price: f64,
    pub listed_price: f64,
    pub is_below_listed: String,
    pub is_in_policy: String,
    pub policy: String,
    pub base_price_after_policy: f64,
    pub gross_margin_rate: f64,
    pub sales_quantity: f64,
    pub pay_amount: f64,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct ActivityPolicyRow {
    pub product_code: String,
    pub platform_activity: String,
    pub start_date: String,
    pub end_date: String,
    pub activity_price: f64,  // 活动后单价
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct OutOfPolicyResult {
    pub file_path: String,
    pub rows: Vec<OutOfPolicyRow>,
    pub total_rows: usize,
    pub load_time_ms: u64,
}

pub fn load_out_of_policy_file(file_path: &str) -> Result<OutOfPolicyResult, String> {
    let start_time = std::time::Instant::now();
    let path = Path::new(file_path);

    let mut workbook: Xlsx<_> =
        open_workbook(path).map_err(|e| format!("无法打开Excel文件: {}", e))?;

    // ================= 读取分析表 =================
    let analysis_sheet = workbook
        .sheet_names()
        .iter()
        .find(|n| n.trim() == "分析")
        .ok_or("未找到分析工作表")?
        .to_string();

    let analysis_range = workbook
        .worksheet_range(&analysis_sheet)
        .map_err(|e| format!("读取分析表失败: {}", e))?;

    let analysis_header = analysis_range.rows().next().ok_or("分析表为空")?;
    let mut analysis_map = HashMap::new();
    for (i, c) in analysis_header.iter().enumerate() {
        if let Data::String(s) = c {
            analysis_map.insert(s.trim().to_string(), i);
        }
    }

    let analysis_rows: Vec<OutOfPolicyRow> = analysis_range
        .rows()
        .skip(1)
        .collect::<Vec<_>>()
        .par_iter()
        .map(|row| OutOfPolicyRow {
            order_date: parse_date_cell(row, &analysis_map, "下单日期"),
            sales_order_no: parse_string_cell(row, &analysis_map, "销售单号"),
            customer_code: parse_string_cell(row, &analysis_map, "客户编码"),
            customer_name: parse_string_cell(row, &analysis_map, "客户名称"),
            product_code: parse_string_cell(row, &analysis_map, "商品编码"),
            generic_name: parse_string_cell(row, &analysis_map, "通用名"),
            sales_price: parse_float_cell(row, &analysis_map, "销售单价/积分"),
            settlement_price: parse_float_cell(row, &analysis_map, "结算单价"),
            listed_price: parse_float_cell(row, &analysis_map, "挂网价"),
            is_below_listed: parse_string_cell(row, &analysis_map, "是否低于挂网"),
            is_in_policy: parse_string_cell(row, &analysis_map, "是否活动政策内"),
            policy: parse_string_cell(row, &analysis_map, "活动政策"),
            base_price_after_policy: parse_float_cell(row, &analysis_map, "活动后底价"),
            gross_margin_rate: parse_float_cell(row, &analysis_map, "毛利率(%)"),
            sales_quantity: parse_float_cell(row, &analysis_map, "销售数量"),
            pay_amount: parse_float_cell(row, &analysis_map, "支付金额"),
        })
        .collect();

    // ================= 读取活动政策表 =================
    let policy_sheet = workbook
        .sheet_names()
        .iter()
        .find(|n| n.trim() == "2025活动政策")
        .ok_or("未找到2025活动政策工作表")?
        .to_string();

    let policy_range = workbook
        .worksheet_range(&policy_sheet)
        .map_err(|e| format!("读取活动政策表失败: {}", e))?;

    let policy_header = policy_range.rows().next().ok_or("活动政策表为空")?;
    let mut policy_map = HashMap::new();
    for (i, c) in policy_header.iter().enumerate() {
        if let Data::String(s) = c {
            let col_name = s.trim().to_string();
            policy_map.insert(col_name.clone(), i);
        }
    }

    let policies: Vec<ActivityPolicyRow> = policy_range
        .rows()
        .skip(1)
        .enumerate()
        .map(|(_row_idx, row)| {
            let product_code = parse_string_cell(row, &policy_map, "商品编码");

            // 直接按索引17获取"平台活动"列的值
            let platform_activity = match row.get(17) {
                Some(Data::String(s)) => s.trim().to_string(),
                Some(Data::Empty) => String::new(),
                Some(cell) => format!("{:?}", cell),
                None => String::new()
            };

            let start_date = parse_date_cell_to_string(row, &policy_map, "开始时间");
            let end_date = parse_date_cell_to_string(row, &policy_map, "结束时间");
            let activity_price = parse_float_cell(row, &policy_map, "活动后单价");

            ActivityPolicyRow {
                product_code,
                platform_activity,
                start_date,
                end_date,
                activity_price,
            }
        })
        .collect();

    // ================= 匹配逻辑 =================
    let mut matched_count = 0;
    let matched_rows: Vec<OutOfPolicyRow> = analysis_rows
        .into_iter()
        .map(|mut row| {
            // 使用新的日期解析函数，支持 Excel 日期数字和字符串格式
            if let Some(od) = parse_date_from_analysis(&row.order_date) {
                if let Some(policy) = find_matching_policy(&policies, &row.product_code, od, row.settlement_price) {
                    row.policy = policy.platform_activity.clone();
                    matched_count += 1;
                }
            }
            row
        })
        .collect();

    println!("总数据行数: {}", matched_rows.len());
    println!("匹配成功的行数: {}", matched_count);
    println!("活动政策数量: {}", policies.len());

    // ================= 修改原Excel文件 =================
    // 创建新的Excel文件，保留原始数据并更新"活动政策"列
    let output_path = file_path.to_string().replace(".xlsx", "_已更新.xlsx");

    let mut workbook = Workbook::new();

    let worksheet = workbook.add_worksheet();

    // 写入表头
    let headers = [
        "下单日期", "销售单号", "客户编码", "客户名称", "商品编码", "通用名",
        "销售单价/积分", "结算单价", "挂网价", "是否低于挂网", "是否活动政策内",
        "活动政策", "活动后底价", "毛利率(%)", "销售数量", "支付金额",
    ];
    for (col, header) in headers.iter().enumerate() {
        worksheet.write_string(0, col as u16, *header)
            .map_err(|e| format!("写入表头失败: {}", e))?;
    }

    // 写入数据
    for (row_idx, row) in matched_rows.iter().enumerate() {
        worksheet.write_string(row_idx as u32 + 1, 0, &row.order_date)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_string(row_idx as u32 + 1, 1, &row.sales_order_no)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_string(row_idx as u32 + 1, 2, &row.customer_code)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_string(row_idx as u32 + 1, 3, &row.customer_name)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_string(row_idx as u32 + 1, 4, &row.product_code)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_string(row_idx as u32 + 1, 5, &row.generic_name)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_number(row_idx as u32 + 1, 6, row.sales_price)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_number(row_idx as u32 + 1, 7, row.settlement_price)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_number(row_idx as u32 + 1, 8, row.listed_price)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_string(row_idx as u32 + 1, 9, &row.is_below_listed)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_string(row_idx as u32 + 1, 10, &row.is_in_policy)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_string(row_idx as u32 + 1, 11, &row.policy)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_number(row_idx as u32 + 1, 12, row.base_price_after_policy)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_number(row_idx as u32 + 1, 13, row.gross_margin_rate)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_number(row_idx as u32 + 1, 14, row.sales_quantity)
            .map_err(|e| format!("写入数据失败: {}", e))?;
        worksheet.write_number(row_idx as u32 + 1, 15, row.pay_amount)
            .map_err(|e| format!("写入数据失败: {}", e))?;
    }

    workbook.save(&output_path)
        .map_err(|e| format!("保存Excel文件失败: {}", e))?;

    Ok(OutOfPolicyResult {
        file_path: output_path,
        total_rows: matched_rows.len(),
        rows: matched_rows,
        load_time_ms: start_time.elapsed().as_millis() as u64,
    })
}

// ==================== 核心辅助函数 ====================
// 将 "分析" 表的日期（可能是数字或字符串）转换为 NaiveDate
fn parse_date_from_analysis(raw: &str) -> Option<NaiveDate> {
    let raw = raw.trim();
    
    // 尝试直接解析为 Excel 日期数字
    if let Ok(excel_date) = raw.parse::<f64>() {
        if excel_date > 0.0 {
            // Excel 日期转换：基准日期 1899-12-30
            let base = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
            let days = excel_date as i64 - 2; // 修正 Excel 闰年错误
            return Some(base + chrono::Duration::days(days));
        }
    }
    
    // 尝试解析为字符串日期格式 "2025/11/4" 或 "2025-11-04"
    let date_part = raw.split_whitespace().next().unwrap_or(raw);
    let sep = if date_part.contains('/') { '/' } else if date_part.contains('-') { '-' } else { return None; };
    let parts: Vec<&str> = date_part.split(sep).collect();
    if parts.len() != 3 { return None; }
    NaiveDate::from_ymd_opt(parts[0].parse().ok()?, parts[1].parse().ok()?, parts[2].parse().ok()?)
}

// 将活动政策表的日期字符串转换为 NaiveDate
fn parse_date(raw: &str) -> Option<NaiveDate> {
    let raw = raw.trim();
    let date_part = raw.split_whitespace().next().unwrap_or(raw);
    let sep = if date_part.contains('/') { '/' } else if date_part.contains('-') { '-' } else { return None; };
    let parts: Vec<&str> = date_part.split(sep).collect();
    if parts.len() != 3 { return None; }
    NaiveDate::from_ymd_opt(parts[0].parse().ok()?, parts[1].parse().ok()?, parts[2].parse().ok()?)
}

fn find_matching_policy<'a>(policies: &'a [ActivityPolicyRow], product_code: &str, order_date: NaiveDate, settlement_price: f64) -> Option<&'a ActivityPolicyRow> {
    let mut price_matched: Option<&'a ActivityPolicyRow> = None;

    for p in policies {
        if p.product_code == product_code {
            if let (Some(s), Some(e)) = (parse_date(&p.start_date), parse_date(&p.end_date)) {
                if order_date >= s && order_date <= e {
                    // 结算单价必须大于活动后单价才算活动
                    if settlement_price > p.activity_price {
                        println!("匹配成功: 商品={}, 日期={}, 结算价={}, 活动价={}, 活动={}",
                            product_code, order_date, settlement_price, p.activity_price, p.platform_activity);
                        price_matched = Some(p);
                        break;
                    } else {
                        println!("价格不匹配: 商品={}, 日期={}, 结算价={}, 活动价={}",
                            product_code, order_date, settlement_price, p.activity_price);
                    }
                }
            }
        }
    }

    // 只返回价格匹配的（结算单价 > 活动后单价）
    price_matched
}

fn parse_string_cell(
    row: &[Data],
    map: &HashMap<String, usize>,
    col: &str,
) -> String {
    let col_trimmed = col.trim();

    // 查找列索引
    let col_idx = match map.get(col_trimmed) {
        Some(&idx) => idx,
        None => return String::new(),
    };

    // 获取单元格数据
    let cell_data = match row.get(col_idx) {
        Some(cell) => cell,
        None => return String::new(),
    };

    // 解析数据
    match cell_data {
        Data::String(s) => s.trim().to_string(),
        Data::Float(f) => f.to_string(),
        Data::Int(i) => i.to_string(),
        Data::Bool(b) => b.to_string(),
        Data::Empty => String::new(),
        Data::DateTime(dt) => dt.to_string(),
        Data::DateTimeIso(s) => s.clone(),
        Data::Error(e) => format!("Error({:?})", e),
        _other => String::new(),
    }
}


fn parse_float_cell(row: &[Data], map: &HashMap<String, usize>, col: &str) -> f64 {
    map.get(col).and_then(|&i| row.get(i)).and_then(|c| match c {
        Data::Float(f) => Some(*f),
        Data::Int(i) => Some(*i as f64),
        Data::String(s) => s.trim().parse().ok(),
        _ => None,
    }).unwrap_or(0.0)
}

fn parse_date_cell(row: &[Data], map: &HashMap<String, usize>, col: &str) -> String {
    map.get(col).and_then(|&i| row.get(i)).map(|c| match c {
        Data::String(s) => s.trim().to_string(),
        Data::Float(f) => excel_date_to_string(*f),
        Data::DateTime(dt) => excel_date_to_string(dt.as_f64()),
        _ => String::new(),
    }).unwrap_or_default()
}

fn parse_date_cell_to_string(row: &[Data], map: &HashMap<String, usize>, col: &str) -> String {
    parse_date_cell(row, map, col)
}

fn excel_date_to_string(v: f64) -> String {
    let base = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
    let d = base + chrono::Duration::days(v as i64); // 修复 Excel 日期偏移
    format!("{}/{}/{}", d.year(), d.month(), d.day())
}
