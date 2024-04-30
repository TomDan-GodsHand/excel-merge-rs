use calamine::{open_workbook, Reader, Xlsx}; // 导入 calamine 模块
use rust_xlsxwriter::{Workbook, XlsxError}; // 导入 rust_xlsxwriter 模块
use std::{
    collections::HashMap,  // 导入 HashMap
    env,                   // 导入环境模块
    error::Error,          // 导入错误模块
    fs,                    // 导入操作模块
    path::{Path, PathBuf}, // 导入引用计数模块
    sync::{Arc, Mutex},    // 导入同步模块
    thread::{self},        // 导入时间相关模块
};
use toml::Value; // 导入 toml 模块

// 定义 SheetContent 结构体
struct SheetContent {
    index: i32,                // 索引
    file_name: String,         // 文件名
    sheet_name: String,        // 表格名
    row_num: i32,              // 行数
    column_num: i32,           // 列数
    content: Vec<Vec<String>>, // 内容
}

// 读取配置文件并返回配置信息
fn read_config(config_path: &Path) -> Result<Value, Box<dyn std::error::Error>> {
    let config_str = fs::read_to_string(config_path)?; // 读取配置文件内容

    let config_value: Value = toml::from_str(&config_str)?; // 解析配置文件内容为 toml::Value

    Ok(config_value) // 返回配置信息
}

// 主函数
fn main() -> Result<(), Box<dyn Error>> {
    let config_path = "config.toml"; // 配置文件路径
    let config_value = read_config(Path::new(config_path))?; // 读取配置文件内容
    let current_dir = env::current_dir()?; // 获取当前目录路径
    println!(
        "Entries modified in the last 24 hours in {:?}:",
        current_dir
    );

    // 读取配置中的 input 部分
    let input_section = config_value.get("input").unwrap().as_table().unwrap();
    let mut folder = input_section.get("folder").unwrap().to_string(); // 获取文件夹路径
    let chcar = folder.remove(0); // 移除首字符
    let chacar = folder.remove(folder.len() - 1); // 移除末尾字符
    let path = PathBuf::from(&folder); // 转换为路径类型
    println!("Input folder: {:?}", &path);

    // 读取配置中的 output 部分
    let output_section = config_value.get("output").unwrap().as_table().unwrap();
    let mut file = output_section.get("file").unwrap().to_string(); // 获取输出文件路径
    let ch = file.remove(0); // 移除首字符
    let chh = file.remove(file.len() - 1); // 移除末尾字符
    println!("Output file: {}", file);

    // 获取文件夹中的文件列表
    let mut files = fs::read_dir(&path)?
        .filter_map(|entry| entry.ok())
        .filter(|entry| entry.path().is_file())
        .map(|entry| entry.path())
        .collect::<Vec<_>>();
   files.sort_by(|a, b| {
        let a_num: i32 = a.file_name().unwrap().to_str().unwrap().split("、").next().unwrap().parse().unwrap();
        let b_num: i32 = b.file_name().unwrap().to_str().unwrap().split("、").next().unwrap().parse().unwrap();
        a_num.cmp(&b_num)
    });
    let mut handles = Vec::new(); // 存放线程句柄的向量
    let mut sheet_contents_hashmap: Arc<Mutex<HashMap<usize, HashMap<usize, SheetContent>>>> =
        Arc::new(Mutex::new(HashMap::new())); // 线程安全的哈希映射
    for (index, file) in files.iter().enumerate() {
        let mut file_path = file.clone();
        let mut file_name = file_path.file_name().unwrap().to_str().unwrap().to_string();
        println!("file{}", file_name);
        let arc_sheet_hashmap = Arc::clone(&sheet_contents_hashmap);
        // 创建线程处理文件
        handles.push(thread::spawn(move || {
            let mut workbook: Xlsx<_> = open_workbook(file_path).unwrap(); // 打开 Excel 文件
            for (worksheet_index, sheet) in workbook.worksheets().iter().enumerate() {
                let mut sheet_range = &sheet.1; // 获取表格范围
                let mut worksheet_name = sheet.0.clone(); // 获取表格名
                let mut content_data: Vec<Vec<String>> = Vec::new(); // 存放表格内容的向量
                let col_num: i32 = sheet_range.get_size().1 as i32; // 列数
                let row_num = sheet_range.get_size().0 as i32; // 行数
                for (row_index, row) in sheet_range.rows().enumerate() {
                    let mut row_vec: Vec<String> = Vec::new(); // 存放行内容的向量
                    for (col_index, cell) in row.iter().enumerate() {
                        let mut cell = cell.to_string(); // 将单元格内容转换为字符串
                        row_vec.push(cell); // 添加到行向量中
                    }
                    content_data.push(row_vec); // 添加到表格内容向量中
                }
                let sheet_content = SheetContent {
                    // 创建 SheetContent 实例
                    index: index as i32,          // 索引
                    file_name: file_name.clone(), // 文件名
                    sheet_name: worksheet_name,   // 表格名
                    row_num: row_num,             // 行数
                    column_num: col_num,          // 列数
                    content: content_data,        // 内容
                };
                println!(
                    "sheet {},{},{}",
                    &sheet_content.sheet_name, &sheet_content.file_name, &sheet_content.index
                );
                let mut sheets_hashmap = arc_sheet_hashmap.lock().unwrap(); // 获取哈希映射锁
                let mut sheet_contents_hashmap = sheets_hashmap
                    .entry(worksheet_index)
                    .or_insert(HashMap::new()); // 获取或插入哈希映射项
                sheet_contents_hashmap.insert(index, sheet_content); // 向哈希映射中插入数据
            }
            drop(workbook); // 释放 Workbook
        }));
    }
    for handle in handles {
        handle.join().unwrap(); // 等待线程结束
    }
    println!("Start output");
    write_excel(file, &sheet_contents_hashmap.lock().unwrap()).expect("output error"); // 输出 Excel 文件
    println!("End output");
    Ok(())
}

// 写入 Excel 文件
pub fn write_excel(
    file: String,
    content: &HashMap<usize, HashMap<usize, SheetContent>>,
) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new(); // 创建 Workbook 实例
    for index in 0..content.len() {
        let sheet_content_hashmap = content.get(&index).unwrap(); // 获取表格内容哈希映射
        let worksheet = workbook.add_worksheet(); // 添加工作表
        match &sheet_content_hashmap.get(&0) {
            Some(sheet_content) => {
                worksheet.set_name(&sheet_content.sheet_name)?; // 设置工作表名
            }
            None => {
                worksheet.set_name("sheet unnamed")?; // 若无工作表名则设置为默认值
            }
        }
        let mut row_offset: u32 = 0; // 行偏移量
        for sheet_index in 0..sheet_content_hashmap.len() {
            let sheet = sheet_content_hashmap.get(&sheet_index).unwrap(); // 获取表格内容
            for (row_index, row_vec) in sheet.content.iter().enumerate() {
                if row_index == 0 {
                    worksheet.write_string(
                        row_index as u32 + row_offset + 1,
                        0,
                        &sheet.file_name,
                    )?; // 写入文件名
                    worksheet.write_string(
                        row_index as u32 + row_offset,
                        0,
                        "序号：".to_string() + &sheet.index.to_string(),
                    )?; // 写入序号
                }
                for (col_index, cell) in row_vec.iter().enumerate() {
                    worksheet.write_string(
                        row_index as u32 + row_offset,
                        (col_index + 1) as u16,
                        cell,
                    )?; // 写入单元格内容
                }
            }

            row_offset += sheet.row_num as u32; // 更新行偏移量
            worksheet.write_string(row_offset + 1, 0, "")?; // 写入空行
            row_offset += 1; // 更新行偏移量
        }
    }
    workbook.save(file)?; // 保存 Workbook
    drop(workbook); // 释放 Workbook
    Ok(())
}
