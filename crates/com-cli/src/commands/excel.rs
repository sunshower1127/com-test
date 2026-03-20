use anyhow::{bail, Result};
use excel_bridge::ExcelApp;

use crate::state::AppState;

pub fn handle(args: &[String], state: &mut AppState) -> Result<()> {
    let action = args.first().map(|s| s.as_str()).unwrap_or("help");

    match action {
        "launch" => {
            if state.excel.is_some() {
                println!("Excel is already running.");
                return Ok(());
            }
            let app = ExcelApp::new()?;
            app.set_visible(true)?;
            state.excel = Some(app);
            println!("[OK] Excel launched.");
        }

        "new" => {
            let app = require_excel(state)?;
            let wb = app.workbooks()?.add()?;
            let sheet = wb.sheet(1)?;
            let name = sheet.name()?;
            state.excel_sheet = Some(sheet);
            state.excel_workbook = Some(wb);
            state.excel_sheet_index = 1;
            println!("[OK] New workbook created. Active sheet: {name}");
        }

        "open" => {
            let path = require_arg(args, 1, "excel open <path>")?;
            let app = require_excel(state)?;
            let wb = app.workbooks()?.open(&path)?;
            let sheet = wb.sheet(1)?;
            let name = sheet.name()?;
            state.excel_sheet = Some(sheet);
            state.excel_workbook = Some(wb);
            state.excel_sheet_index = 1;
            println!("[OK] Opened. Active sheet: {name}");
        }

        "sheet" => {
            let index: i32 = args.get(1).and_then(|s| s.parse().ok()).unwrap_or(1);
            let wb = require_workbook(state)?;
            let sheet = wb.sheet(index)?;
            let name = sheet.name()?;
            state.excel_sheet = Some(sheet);
            state.excel_sheet_index = index;
            println!("[OK] Active sheet: {name} (#{index})");
        }

        "cell" => {
            let addr = require_arg(args, 1, "excel cell <A1>")?;
            let sheet = require_sheet(state)?;
            let val = sheet.range(&addr)?.value()?;
            println!("  {addr} = {val}");
        }

        "set" => {
            let addr = require_arg(args, 1, "excel set <A1> <value>")?;
            let value_str = args.get(2..).map(|s| s.join(" ")).unwrap_or_default();
            if value_str.is_empty() {
                bail!("Usage: excel set <A1> <value>");
            }
            let sheet = require_sheet(state)?;
            let value = parse_value(&value_str);
            sheet.range(&addr)?.set_value(value)?;
            println!("[OK] {addr} = {value_str}");
        }

        "formula" => {
            let addr = require_arg(args, 1, "excel formula <A1> <expr>")?;
            let expr = require_arg(args, 2, "excel formula <A1> <expr>")?;
            let sheet = require_sheet(state)?;
            sheet.range(&addr)?.set_formula(&expr)?;
            println!("[OK] {addr} formula = {expr}");
        }

        "getformula" => {
            let addr = require_arg(args, 1, "excel getformula <A1>")?;
            let sheet = require_sheet(state)?;
            let f = sheet.range(&addr)?.formula()?;
            println!("  {addr} formula = {f}");
        }

        "save" => {
            let path = require_arg(args, 1, "excel save <path>")?;
            let wb = require_workbook(state)?;
            wb.save_as(&path)?;
            println!("[OK] Saved to {path}");
        }

        "close" => {
            let wb = require_workbook(state)?;
            wb.close()?;
            state.excel_sheet = None;
            state.excel_workbook = None;
            state.excel_sheet_index = 0;
            println!("[OK] Workbook closed.");
        }

        "quit" => {
            if let Some(app) = state.excel.take() {
                let _ = app.set_display_alerts(false);
                let _ = app.quit();
                std::mem::forget(app); // Drop에서 중복 Quit 방지
            }
            state.clear_excel();
            println!("[OK] Excel quit.");
        }

        _ => println!("Unknown excel command: {action} (type 'help')"),
    }

    Ok(())
}

fn require_excel(state: &AppState) -> Result<&ExcelApp> {
    state.excel.as_ref().ok_or_else(|| anyhow::anyhow!("Excel is not running. Use 'excel launch' first."))
}

fn require_workbook(state: &AppState) -> Result<&excel_bridge::Workbook> {
    state.excel_workbook.as_ref().ok_or_else(|| anyhow::anyhow!("No workbook open. Use 'excel new' or 'excel open'."))
}

fn require_sheet(state: &AppState) -> Result<&excel_bridge::Worksheet> {
    state.excel_sheet.as_ref().ok_or_else(|| anyhow::anyhow!("No sheet selected. Use 'excel sheet [n]'."))
}

fn require_arg(args: &[String], index: usize, usage: &str) -> Result<String> {
    args.get(index)
        .cloned()
        .ok_or_else(|| anyhow::anyhow!("Usage: {usage}"))
}

fn parse_value(s: &str) -> com_core::Variant {
    if let Ok(n) = s.parse::<i32>() {
        return com_core::Variant::I32(n);
    }
    if let Ok(n) = s.parse::<f64>() {
        return com_core::Variant::F64(n);
    }
    match s {
        "true" => com_core::Variant::Bool(true),
        "false" => com_core::Variant::Bool(false),
        _ => com_core::Variant::String(s.to_string()),
    }
}
