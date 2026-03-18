use com_core::ComRuntime;
use excel_bridge::ExcelApp;

fn main() -> windows::core::Result<()> {
    // 1. COM STA 초기화
    let _com = ComRuntime::init()?;
    println!("[OK] COM initialized");

    // 2. Excel 실행
    let excel = ExcelApp::new()?;
    excel.set_visible(true)?;
    excel.set_display_alerts(false)?;
    println!("[OK] Excel launched");

    // 3. 새 워크북 생성
    let wb = excel.workbooks()?.add()?;
    let sheet = wb.sheet(1)?;
    println!("[OK] Workbook created, sheet: {}", sheet.name()?);

    // 4. 셀 쓰기
    sheet.cell(1, 1)?.set_value("Hello")?;
    sheet.cell(1, 2)?.set_value("World")?;
    sheet.cell(2, 1)?.set_value(42)?;
    sheet.cell(2, 2)?.set_formula("=A2*2")?;
    println!("[OK] Cells written");

    // 5. 셀 읽기
    let a1 = sheet.cell(1, 1)?.value()?;
    let b1 = sheet.cell(1, 2)?.value()?;
    let a2 = sheet.cell(2, 1)?.value()?;
    let b2 = sheet.cell(2, 2)?.value()?;
    println!("A1={a1}, B1={b1}, A2={a2}, B2={b2}");

    // 6. 저장
    let save_path = std::env::temp_dir().join("com_test_hello.xlsx");
    let save_path_str = save_path.to_string_lossy().to_string();
    wb.save_as(&save_path_str)?;
    println!("[OK] Saved to {save_path_str}");

    // 7. 정리 (Drop이 자동으로 Quit 호출)
    wb.close()?;
    println!("[OK] Done!");

    Ok(())
}
