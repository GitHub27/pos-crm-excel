// Learn more about Tauri commands at https://tauri.app/develop/calling-rust/
use calamine::{Reader, Xlsx, open_workbook, DataType};
use xlsxwriter::*;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

#[derive(Serialize, Deserialize)]
struct ProcessResult {
    success: bool,
    message: String,
    output_path: Option<String>,
}

#[tauri::command]
fn greet(name: &str) -> String {
    format!("Hello, {}! You've been greeted from Rust!", name)
}

#[tauri::command]
async fn process_excel(input_path: String, output_path: String, column_name: String, create_date: String, update_date: String) -> Result<ProcessResult, String> {
    // 调用核心处理逻辑，如果是错误（如歧义），直接返回带错误信息的 Result
    match process_excel_file(&input_path, &output_path, &column_name, &create_date, &update_date).await {
        Ok(_) => Ok(ProcessResult {
            success: true,
            message: "Excel文件处理成功".to_string(),
            output_path: Some(output_path),
        }),
        Err(e) => Ok(ProcessResult {
            success: false,
            message: e.to_string(), // 直接展示内部抛出的具体错误信息（如“第x行数据歧义”）
            output_path: None,
        }),
    }
}

/// 预检查 Excel 数据中的地址歧义，不生成输出文件。
#[tauri::command]
async fn precheck_excel(input_path: String, column_name: String) -> Result<ProcessResult, String> {
    match precheck_excel_file(&input_path, &column_name) {
        Ok(_) => Ok(ProcessResult {
            success: true,
            message: "预检查通过".to_string(),
            output_path: None,
        }),
        Err(e) => Ok(ProcessResult {
            success: false,
            message: e.to_string(),
            output_path: None,
        }),
    }
}

/// 执行地址歧义检查；若存在“开发区/高新区”且未标明城市则返回错误。
fn precheck_excel_file(input_path: &str, column_name: &str) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook(input_path)?;
    let sheet_names = workbook.sheet_names().to_owned();

    if sheet_names.is_empty() {
        return Err("Excel文件中没有工作表".into());
    }

    let range = workbook.worksheet_range(&sheet_names[0])
        .map_err(|e| format!("无法读取工作表: {}", e))?;

    let mut target_col_index: Option<usize> = None;
    if let Some(first_row) = range.rows().next() {
        for (col_idx, cell) in first_row.iter().enumerate() {
            let cell_str = match cell {
                calamine::Data::String(s) => s.as_str(),
                _ => "",
            };
            if cell_str == column_name {
                target_col_index = Some(col_idx);
                break;
            }
        }
    }

    let col_index = target_col_index.ok_or(format!("未找到列名为'{}'的列", column_name))?;

    for (row_idx, row) in range.rows().enumerate() {
        if row_idx == 0 {
            continue;
        }
        if let Some(cell) = row.get(col_index) {
            let cell_value = match cell {
                calamine::Data::Empty => "",
                calamine::Data::String(s) => s.as_str(),
                _ => "",
            };

            if cell_value.contains("高新区") || cell_value.contains("开发区") {
                let has_xuzhou = cell_value.contains("徐州");
                let has_suqian = cell_value.contains("宿迁");
                let has_jining = cell_value.contains("济宁");

                if !has_xuzhou && !has_suqian && !has_jining {
                    return Err(format!("第 {} 行的数据（{}）地址存在歧义，请补充具体城市（如徐州、宿迁或济宁）", row_idx + 1, cell_value).into());
                }
            }
        }
    }

    Ok(())
}

async fn process_excel_file(input_path: &str, output_path: &str, column_name: &str, create_date: &str, update_date: &str) -> Result<(), Box<dyn std::error::Error>> {
    // 创建映射关系
    let mut mapping = vec![
        ("鼓楼", ("江苏省徐州市鼓楼区", "320000,320300,320302")),
        ("云龙", ("江苏省徐州市云龙区", "320000,320300,320303")),
        ("泉山", ("江苏省徐州市泉山区", "320000,320300,320311")),
        ("铜山", ("江苏省徐州市铜山区", "320000,320300,320312")),
        ("丰县", ("江苏省徐州市丰县", "320000,320300,320321")),
        ("睢宁", ("江苏省徐州市睢宁县", "320000,320300,320324")),
        ("新沂", ("江苏省徐州市新沂市", "320000,320300,320381")),
        ("邳州", ("江苏省徐州市邳州市", "320000,320300,320382")),
        ("工业园", ("江苏省徐州市工业园区", "320000,320300,320391")),
        ("工业区", ("江苏省徐州市工业园区", "320000,320300,320391")),
        ("贾汪", ("江苏省徐州市贾汪区", "320000,320300,320305")),
        ("沛县", ("江苏省徐州市沛县", "320000,320300,320322")),
        // 宿迁区域
        ("宿城", ("江苏省宿迁市宿城区", "320000,321300,321302")),
        ("宿豫", ("江苏省宿迁市宿豫区", "320000,321300,321311")),
        ("沭阳", ("江苏省宿迁市沭阳县", "320000,321300,321322")),
        ("泗阳", ("江苏省宿迁市泗阳县", "320000,321300,321323")),
        ("泗洪", ("江苏省宿迁市泗洪县", "320000,321300,321324")),
        ("宿迁经济区", ("江苏省宿迁市宿迁经济技术开发区", "320000,321300,321371")),
        // 济宁区域
        ("任城", ("山东省济宁市任城区", "370000,370800,370811")),
        ("兖州", ("山东省济宁市兖州区", "370000,370800,370812")),
        ("微山", ("山东省济宁市微山县", "370000,370800,370826")),
        ("鱼台", ("山东省济宁市鱼台县", "370000,370800,370827")),
        ("金乡", ("山东省济宁市金乡县", "370000,370800,370828")),
        ("嘉祥", ("山东省济宁市嘉祥县", "370000,370800,370829")),
        ("汶上", ("山东省济宁市汶上县", "370000,370800,370830")),
        ("泗水", ("山东省济宁市泗水县", "370000,370800,370831")),
        ("梁山", ("山东省济宁市梁山县", "370000,370800,370832")),
        ("曲阜", ("山东省济宁市曲阜市", "370000,370800,370881")),
        ("邹城", ("山东省济宁市邹城市", "370000,370800,370883")),
    ];

    // 按关键字长度降序排序，保证长字符串（如"宿迁开发区"）优先匹配，避免被短字符串（如"开发区"）截胡
    mapping.sort_by(|a, b| b.0.chars().count().cmp(&a.0.chars().count()));

    
    // 读取Excel文件
    let mut workbook: Xlsx<_> = open_workbook(input_path)?;
    let sheet_names = workbook.sheet_names().to_owned();
    
    if sheet_names.is_empty() {
        return Err("Excel文件中没有工作表".into());
    }
    
    // 读取第一个工作表
    let range = workbook.worksheet_range(&sheet_names[0])
        .map_err(|e| format!("无法读取工作表: {}", e))?;
    
    // 查找目标列的索引
    let mut target_col_index: Option<usize> = None;
    if let Some(first_row) = range.rows().next() {
        for (col_idx, cell) in first_row.iter().enumerate() {
            let cell_str = match cell {
                calamine::Data::String(s) => s.as_str(),
                _ => "",
            };
            if cell_str == column_name {
                target_col_index = Some(col_idx);
                break;
            }
        }
    }
    
    let col_index = target_col_index.ok_or(format!("未找到列名为'{}'的列", column_name))?;

    // 第一遍遍历：检查所有数据，如果有歧义地址，直接报错退出，不生成文件
    for (row_idx, row) in range.rows().enumerate() {
        if row_idx == 0 {
            continue; // 跳过标题行
        }
        if let Some(cell) = row.get(col_index) {
            let cell_value = match cell {
                calamine::Data::Empty => "",
                calamine::Data::String(s) => s.as_str(),
                _ => "",
            };

            if cell_value.contains("高新区") || cell_value.contains("开发区") {
                let has_xuzhou = cell_value.contains("徐州");
                let has_suqian = cell_value.contains("宿迁");
                let has_jining = cell_value.contains("济宁");

                if !has_xuzhou && !has_suqian && !has_jining {
                    // +1 是因为 row_idx 是从 0 开始的，加上 Excel 真实的行号应该更直观 (+1 标题行, +1 row_idx)
                    return Err(format!("第 {} 行的数据（{}）地址存在歧义，请补充具体城市（如徐州、宿迁或济宁）", row_idx + 1, cell_value).into());
                }
            }
        }
    }
    
    // 创建新的Excel文件
    let workbook_writer = Workbook::new(output_path)?;
    let mut worksheet = workbook_writer.add_worksheet(Some(&sheet_names[0]))?;
    
    // 处理数据
    for (row_idx, row) in range.rows().enumerate() {
        let mut row_code_value = "".to_string();

        for (col_idx, cell) in row.iter().enumerate() {
            let mut cell_value = match cell {
                calamine::Data::Empty => "".to_string(),
                calamine::Data::String(s) => s.clone(),
                calamine::Data::Float(f) => f.to_string(),
                calamine::Data::Int(i) => i.to_string(),
                calamine::Data::Bool(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
                _ => "".to_string(),
            };
            
            // 如果是目标列且不是标题行，检查是否需要替换
            if col_idx == col_index && row_idx > 0 {
                if cell_value.contains("高新区") || cell_value.contains("开发区") {
                    let has_xuzhou = cell_value.contains("徐州");
                    let has_suqian = cell_value.contains("宿迁");
                    let has_jining = cell_value.contains("济宁");

                    if has_xuzhou {
                        cell_value = "江苏省徐州市徐州经济技术开发区".to_string();
                        row_code_value = "320000,320300,320371".to_string();
                    } else if has_suqian {
                        cell_value = "江苏省宿迁市宿迁经济技术开发区".to_string();
                        row_code_value = "320000,321300,321371".to_string();
                    } else if has_jining {
                        cell_value = "山东省济宁市济宁高新技术产业开发区".to_string();
                        row_code_value = "370000,370800,370871".to_string();
                    }
                } else {
                    for (key, (full_name, code)) in &mapping {
                        if cell_value.contains(key) {
                            cell_value = full_name.to_string();
                            row_code_value = code.to_string();
                            break;
                        }
                    }
                }
            }
            
            worksheet.write_string(row_idx as u32, col_idx as u16, &cell_value, None)?;
        }
        
        // 如果是标题行，添加"省市区编码"、"创建日期"和"更新日期"列标题
        if row_idx == 0 {
            let code_col_index = row.len() as u16;
            let create_date_col_index = code_col_index + 1;
            let update_date_col_index = code_col_index + 2;
            
            worksheet.write_string(0, code_col_index, "省市区编码", None)?;
            worksheet.write_string(0, create_date_col_index, "创建日期", None)?;
            worksheet.write_string(0, update_date_col_index, "更新日期", None)?;
        } else {
            // 为数据行添加对应的编码和日期
            let code_col_index = row.len() as u16;
            let create_date_col_index = code_col_index + 1;
            let update_date_col_index = code_col_index + 2;
            
            worksheet.write_string(row_idx as u32, code_col_index, &row_code_value, None)?;
            worksheet.write_string(row_idx as u32, create_date_col_index, create_date, None)?;
            worksheet.write_string(row_idx as u32, update_date_col_index, update_date, None)?;
        }
    }
    
    workbook_writer.close()?;
    Ok(())
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![greet, process_excel, precheck_excel])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
