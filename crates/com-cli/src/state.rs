use com_core::DispatchObject;
use excel_bridge::{ExcelApp, Workbook, Worksheet};
use hwp_bridge::HwpApp;

/// CLI 세션의 전체 상태
/// Drop 순서: 의존 객체 → 앱 (선언 순서대로 drop)
pub struct AppState {
    // Excel
    pub excel_sheet: Option<Worksheet>,
    pub excel_workbook: Option<Workbook>,
    pub excel: Option<ExcelApp>,
    pub excel_sheet_index: i32,

    // HWP
    pub hwp: Option<HwpApp>,
    pub hwp_file: Option<String>,

    // Raw dispatch
    pub raw_target: Option<DispatchObject>,
    pub last_result: Option<DispatchObject>,
}

impl AppState {
    pub fn new() -> Self {
        Self {
            excel_sheet: None,
            excel_workbook: None,
            excel: None,
            excel_sheet_index: 0,
            hwp: None,
            hwp_file: None,
            raw_target: None,
            last_result: None,
        }
    }

    pub fn print_status(&self) {
        println!("=== Status ===");

        // Excel
        if self.excel.is_some() {
            print!("Excel: running");
            if self.excel_workbook.is_some() {
                if let Some(ref sheet) = self.excel_sheet {
                    if let Ok(name) = sheet.name() {
                        print!(" | sheet: {} (#{}) ", name, self.excel_sheet_index);
                    }
                }
            } else {
                print!(" | no workbook");
            }
            println!();
        } else {
            println!("Excel: not running");
        }

        // HWP
        if self.hwp.is_some() {
            print!("HWP: running");
            if let Some(ref f) = self.hwp_file {
                print!(" | file: {f}");
            }
            println!();
        } else {
            println!("HWP: not running");
        }

        // Raw
        if self.raw_target.is_some() {
            println!("Raw target: set");
        }
    }

    /// Excel 종료 시 관련 상태 정리
    pub fn clear_excel(&mut self) {
        self.excel_sheet = None;
        self.excel_workbook = None;
        self.excel = None;
        self.excel_sheet_index = 0;
    }

    /// HWP 종료 시 관련 상태 정리
    pub fn clear_hwp(&mut self) {
        self.hwp = None;
        self.hwp_file = None;
    }
}
