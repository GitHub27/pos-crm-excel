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
async fn process_excel(input_path: String, output_path: String, column_name: String) -> Result<ProcessResult, String> {
    match process_excel_file(&input_path, &output_path, &column_name).await {
        Ok(_) => Ok(ProcessResult {
            success: true,
            message: "Excel文件处理成功".to_string(),
            output_path: Some(output_path),
        }),
        Err(e) => Ok(ProcessResult {
            success: false,
            message: format!("处理失败: {}", e),
            output_path: None,
        }),
    }
}

async fn process_excel_file(input_path: &str, output_path: &str, column_name: &str) -> Result<(), Box<dyn std::error::Error>> {
    // 创建映射关系
    let mut mapping = HashMap::new();
    mapping.insert("鼓楼", ("江苏省徐州市鼓楼区", "320000,320300,320302"));
    mapping.insert("云龙", ("江苏省徐州市云龙区", "320000,320300,320303"));
    mapping.insert("泉山", ("江苏省徐州市泉山区", "320000,320300,320311"));
    mapping.insert("铜山", ("江苏省徐州市铜山区", "320000,320300,320312"));
    mapping.insert("丰县", ("江苏省徐州市丰县", "320000,320300,320321"));
    mapping.insert("睢宁", ("江苏省徐州市睢宁县", "320000,320300,320324"));
    mapping.insert("开发区", ("江苏省徐州市徐州经济技术开发区", "320000,320300,320371"));
    mapping.insert("新沂", ("江苏省徐州市新沂市", "320000,320300,320381"));
    mapping.insert("邳州", ("江苏省徐州市邳州市", "320000,320300,320382"));
    mapping.insert("工业园", ("江苏省徐州市工业园区", "320000,320300,320391"));
    mapping.insert("贾汪", ("江苏省徐州市贾汪区", "320000,320300,320305"));
    mapping.insert("沛县", ("江苏省徐州市沛县", "320000,320300,320322"));
    
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
    
    // 创建新的Excel文件
    let workbook_writer = Workbook::new(output_path)?;
    let mut worksheet = workbook_writer.add_worksheet(Some(&sheet_names[0]))?;
    
    // 处理数据
    for (row_idx, row) in range.rows().enumerate() {
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
                for (key, (full_name, _)) in &mapping {
                    if cell_value.contains(key) {
                        cell_value = full_name.to_string();
                        break;
                    }
                }
            }
            
            worksheet.write_string(row_idx as u32, col_idx as u16, &cell_value, None)?;
        }
        
        // 如果是标题行，添加"省市区编码"列标题
        if row_idx == 0 {
            let code_col_index = row.len() as u16;
            worksheet.write_string(0, code_col_index, "省市区编码", None)?;
        } else {
            // 为数据行添加对应的编码
            let code_col_index = row.len() as u16;
            let mut code_value = "".to_string();
            
            if let Some(cell) = row.get(col_index) {
                let cell_str = match cell {
                    calamine::Data::String(s) => s.as_str(),
                    _ => "",
                };
                
                for (key, (_, code)) in &mapping {
                    if cell_str.contains(key) {
                        code_value = code.to_string();
                        break;
                    }
                }
            }
            
            worksheet.write_string(row_idx as u32, code_col_index, &code_value, None)?;
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
        .invoke_handler(tauri::generate_handler![greet, process_excel])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
