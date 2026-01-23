use calamine::{open_workbook, Reader, Xlsx, Data, DataType};
use serde::{Deserialize, Serialize};
use std::path::Path;
use rayon::prelude::*;
use chrono::Datelike;

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct OutOfPolicyRow {
    pub order_date: String,          // 下单日期
    pub sales_order_no: String,      // 销售单号
    pub customer_code: String,       // 客户编码
    pub customer_name: String,       // 客户名称
    pub product_code: String,        // 商品编码
    pub generic_name: String,        // 通用名
    pub sales_price: f64,            // 销售单价/积分
    pub settlement_price: f64,       // 结算单价
    pub listed_price: f64,           // 挂网价
    pub is_below_listed: String,     // 是否低于挂网
    pub is_in_policy: String,        // 是否活动政策内
    pub policy: String,              // 活动政策
    pub base_price_after_policy: f64, // 活动后底价
    pub gross_margin_rate: f64,      // 毛利率(%)
    pub sales_quantity: f64,         // 销售数量
    pub pay_amount: f64,             // 支付金额
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct ActivityPolicyRow {
    pub year: String,               // 年
    pub month: String,              // 月
    pub product_type: String,        // 商品类型
    pub plan_dept: String,          // 策划部门
    pub cost_source: String,         // 费用来源
    pub activity: String,            // 活动
    pub activity_type: String,       // 活动类型
    pub activity_form: String,       // 活动形式
    pub product_grade: String,       // 商品分级
    pub serial_no: String,           // 序号
    pub product_code: String,        // 商品编码
    pub generic_name: String,        // 通用名
    pub product_name: String,        // 商品名
    pub spec: String,               // 规格
    pub manufacturer: String,        // 生产企业
    pub purchase_price: f64,         // 进价
    pub listed_price: f64,          // 挂网价
    pub platform_activity: String,   // 平台活动
    pub price_after_activity: f64,  // 活动后单价
    pub margin_after_activity: f64,  // 活动后毛利率
    pub start_date: String,         // 开始时间
    pub end_date: String,           // 结束时间
    pub remark: String,             // 备注
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct OutOfPolicyResult {
    pub file_path: String,
    pub rows: Vec<OutOfPolicyRow>,
    pub total_rows: usize,
    pub load_time_ms: u64,
}

pub fn load_out_of_policy_file(
    file_path: &str,
) -> Result<OutOfPolicyResult, String> {
    let start_time = std::time::Instant::now();
    let path = Path::new(file_path);

    let mut workbook: Xlsx<_> = open_workbook(path)
        .map_err(|e| format!("无法打开Excel文件: {}", e))?;

    // 读取"分析"工作表
    let analysis_sheet_name = workbook.sheet_names()
        .iter()
        .find(|name| name.trim() == "分析")
        .ok_or("Excel文件中未找到名为'分析'的工作表")?
        .to_string();

    let analysis_range = workbook.worksheet_range(&analysis_sheet_name)
        .map_err(|e| format!("读取'分析'工作表数据失败: {}", e))?;

    // 读取"2025活动政策"工作表
    let activity_sheet_name = workbook.sheet_names()
        .iter()
        .find(|name| name.trim() == "2025活动政策")
        .ok_or("Excel文件中未找到名为'2025活动政策'的工作表")?
        .to_string();

    let activity_range = workbook.worksheet_range(&activity_sheet_name)
        .map_err(|e| format!("读取'2025活动政策'工作表数据失败: {}", e))?;

    // 读取活动政策数据
    let activity_header = activity_range.rows().next()
        .ok_or("'2025活动政策'工作表为空".to_string())?;

    let mut activity_header_map: std::collections::HashMap<String, usize> = std::collections::HashMap::new();
    for (i, cell) in activity_header.iter().enumerate() {
        if let Data::String(name) = cell {
            activity_header_map.insert(name.trim().to_string(), i);
        }
    }

    let activity_required_columns = [
        "年", "月", "商品类型", "策划部门", "费用来源", "活动", "活动类型",
        "活动形式", "商品分级", "序号", "商品编码", "通用名", "商品名", "规格",
        "生产企业", "进价", "挂网价", "平台活动", "活动后单价", "活动后毛利率",
        "开始时间", "结束时间", "备注",
    ];

    for col in &activity_required_columns {
        if !activity_header_map.contains_key(*col) {
            return Err(format!("'2025活动政策'工作表缺少必需列: {}", col));
        }
    }

    let activity_data_rows: Vec<&[Data]> = activity_range.rows().skip(1).collect();
    let activity_policies: Vec<ActivityPolicyRow> = activity_data_rows
        .iter()
        .map(|row| ActivityPolicyRow {
            year: parse_string_cell(row, &activity_header_map, "年"),
            month: parse_string_cell(row, &activity_header_map, "月"),
            product_type: parse_string_cell(row, &activity_header_map, "商品类型"),
            plan_dept: parse_string_cell(row, &activity_header_map, "策划部门"),
            cost_source: parse_string_cell(row, &activity_header_map, "费用来源"),
            activity: parse_string_cell(row, &activity_header_map, "活动"),
            activity_type: parse_string_cell(row, &activity_header_map, "活动类型"),
            activity_form: parse_string_cell(row, &activity_header_map, "活动形式"),
            product_grade: parse_string_cell(row, &activity_header_map, "商品分级"),
            serial_no: parse_string_cell(row, &activity_header_map, "序号"),
            product_code: parse_string_cell(row, &activity_header_map, "商品编码"),
            generic_name: parse_string_cell(row, &activity_header_map, "通用名"),
            product_name: parse_string_cell(row, &activity_header_map, "商品名"),
            spec: parse_string_cell(row, &activity_header_map, "规格"),
            manufacturer: parse_string_cell(row, &activity_header_map, "生产企业"),
            purchase_price: parse_float_cell(row, &activity_header_map, "进价"),
            listed_price: parse_float_cell(row, &activity_header_map, "挂网价"),
            platform_activity: parse_string_cell(row, &activity_header_map, "平台活动"),
            price_after_activity: parse_float_cell(row, &activity_header_map, "活动后单价"),
            margin_after_activity: parse_float_cell(row, &activity_header_map, "活动后毛利率"),
            start_date: parse_string_cell(row, &activity_header_map, "开始时间"),
            end_date: parse_string_cell(row, &activity_header_map, "结束时间"),
            remark: parse_string_cell(row, &activity_header_map, "备注"),
        })
        .collect();

    // 简单的列索引映射 (name -> index)
    let mut header_map: std::collections::HashMap<String, usize> = std::collections::HashMap::new();
    
    // 读取分析表头
    let header = analysis_range.rows().next()
        .ok_or("'分析'工作表为空".to_string())?;
    
    for (i, cell) in header.iter().enumerate() {
        if let Data::String(name) = cell {
            header_map.insert(name.trim().to_string(), i);
        }
    }

    // 验证必需列是否存在
    let required_columns = [
        "下单日期", "销售单号", "客户编码", "客户名称", "商品编码", "通用名",
        "销售单价/积分", "结算单价", "挂网价", "是否低于挂网",
        "是否活动政策内", "活动政策", "活动后底价", "毛利率(%)",
        "销售数量", "支付金额",
    ];

    for col in &required_columns {
        if !header_map.contains_key(*col) {
            return Err(format!("缺少必需列: {}", col));
        }
    }

    // 使用并行处理解析数据行 - 提升性能
    let data_rows: Vec<&[Data]> = analysis_range.rows().skip(1).collect();

    let rows: Vec<OutOfPolicyRow> = data_rows
        .par_iter()  // 并行迭代
        .map(|row| OutOfPolicyRow {
            order_date: parse_date_cell(row, &header_map, "下单日期"),
            sales_order_no: parse_string_cell(row, &header_map, "销售单号"),
            customer_code: parse_string_cell(row, &header_map, "客户编码"),
            customer_name: parse_string_cell(row, &header_map, "客户名称"),
            product_code: parse_string_cell(row, &header_map, "商品编码"),
            generic_name: parse_string_cell(row, &header_map, "通用名"),
            sales_price: parse_float_cell(row, &header_map, "销售单价/积分"),
            settlement_price: parse_float_cell(row, &header_map, "结算单价"),
            listed_price: parse_float_cell(row, &header_map, "挂网价"),
            is_below_listed: parse_string_cell(row, &header_map, "是否低于挂网"),
            is_in_policy: parse_string_cell(row, &header_map, "是否活动政策内"),
            policy: parse_string_cell(row, &header_map, "活动政策"),
            base_price_after_policy: parse_float_cell(row, &header_map, "活动后底价"),
            gross_margin_rate: parse_float_cell(row, &header_map, "毛利率(%)"),
            sales_quantity: parse_float_cell(row, &header_map, "销售数量"),
            pay_amount: parse_float_cell(row, &header_map, "支付金额"),
        })
        .collect();

    // 根据商品编码和下单日期匹配活动政策
    let matched_rows: Vec<OutOfPolicyRow> = rows
        .into_iter()
        .map(|mut row| {
            // 解析下单日期 (格式: 2025/11/4)
            let order_date = parse_date(&row.order_date);
            if let Some(od) = order_date {
                // 查找匹配的活动政策
                if let Some(policy) = find_matching_policy(&activity_policies, &row.product_code, od) {
                    // 将平台活动值更新到活动政策列
                    row.policy = policy.platform_activity.clone();
                }
            }
            row
        })
        .collect();

    Ok(OutOfPolicyResult {
        file_path: file_path.to_string(),
        total_rows: matched_rows.len(),
        rows: matched_rows,
        load_time_ms: start_time.elapsed().as_millis() as u64,
    })
}

/// 解析日期字符串 (格式: 2025/11/4)
fn parse_date(date_str: &str) -> Option<chrono::NaiveDate> {
    let parts: Vec<&str> = date_str.split('/').collect();
    if parts.len() >= 3 {
        let year: i32 = parts.get(0)?.parse().ok()?;
        let month: u32 = parts.get(1)?.parse().ok()?;
        let day: u32 = parts.get(2)?.parse().ok()?;
        chrono::NaiveDate::from_ymd_opt(year, month, day)
    } else {
        None
    }
}

/// 查找匹配的活动政策
fn find_matching_policy<'a>(
    policies: &'a [ActivityPolicyRow],
    product_code: &str,
    order_date: chrono::NaiveDate,
) -> Option<&'a ActivityPolicyRow> {
    policies.iter().find(|policy| {
        policy.product_code == product_code && is_date_in_range(&policy.start_date, &policy.end_date, order_date)
    })
}

/// 检查日期是否在范围内
fn is_date_in_range(start_str: &str, end_str: &str, target_date: chrono::NaiveDate) -> bool {
    if let (Some(start), Some(end)) = (parse_date(start_str), parse_date(end_str)) {
        target_date >= start && target_date <= end
    } else {
        false
    }
}

/// 辅助函数：解析字符串单元格
fn parse_string_cell(row: &[Data], header_map: &std::collections::HashMap<String, usize>, col_name: &str) -> String {
    header_map
        .get(col_name)
        .and_then(|&idx| row.get(idx))
        .map(|cell| match cell {
            Data::String(s) => s.trim().to_string(),
            Data::Float(f) => f.to_string(),
            Data::Int(i) => i.to_string(),
            Data::Bool(b) => b.to_string(),
            Data::DateTime(dt) => dt.to_string(),
            Data::DateTimeIso(s) => s.clone(),
            Data::DurationIso(s) => s.clone(),
            _ => String::new(),
        })
        .unwrap_or_default()
}

/// 辅助函数：解析日期单元格（处理 Excel 日期数字）
fn parse_date_cell(row: &[Data], header_map: &std::collections::HashMap<String, usize>, col_name: &str) -> String {
    header_map
        .get(col_name)
        .and_then(|&idx| row.get(idx))
        .map(|cell| {
            match cell {
                Data::String(s) => s.trim().to_string(),
                Data::Float(f) => {
                    // Excel 日期是以数字存储的，需要转换
                    // Excel 基准日期是 1900-01-01（但实际上 1900-01-01 是第1天）
                    // 使用 chrono 转换
                    excel_date_to_string(*f)
                }
                Data::Int(i) => i.to_string(),
                Data::DateTime(dt) => dt.to_string(),
                Data::DateTimeIso(s) => s.clone(),
                _ => String::new(),
            }
        })
        .unwrap_or_default()
}

/// 将 Excel 日期数字转换为日期字符串 (格式: 2025/11/4)
fn excel_date_to_string(excel_date: f64) -> String {
    // 尝试 1900 日期系统（默认）
    // Excel 日期 1 = 1900-01-01，但 Excel 认为存在 1900-02-29（不存在），所以需要减去 2
    let base_date_1900 = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
    let days_since_1899 = excel_date as i32 - 2;

    // 检查是否为合理的日期（1900 系统）
    if excel_date < 1.0 {
        return String::new();
    }

    // 使用 1899-12-30 作为基准日期（修正后的 1900 系统）
    let result_date = base_date_1900 + chrono::Duration::days(days_since_1899 as i64);

    format!(
        "{}/{}/{}",
        result_date.year(),
        result_date.month(),
        result_date.day()
    )
}

/// 辅助函数：解析浮点数单元格
fn parse_float_cell(row: &[Data], header_map: &std::collections::HashMap<String, usize>, col_name: &str) -> f64 {
    header_map
        .get(col_name)
        .and_then(|&idx| row.get(idx))
        .and_then(|cell| match cell {
            Data::Float(f) => Some(*f),
            Data::Int(i) => Some(*i as f64),
            Data::String(s) => s.trim().parse().ok(),
            _ => None,
        })
        .unwrap_or(0.0)
}

