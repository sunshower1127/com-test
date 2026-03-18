use com_core::ComRuntime;
use hwp_bridge::{HwpApp, HwpDetector};

fn main() -> windows::core::Result<()> {
    // 1. 진단
    HwpDetector::diagnose();
    println!();

    // 2. COM 초기화
    let _com = ComRuntime::init()?;
    println!("[OK] COM initialized");

    // 3. HWP 실행
    let hwp = HwpApp::new()?;
    hwp.set_visible(true)?;
    println!("[OK] HWP launched");

    // 4. 보안 모듈 등록 시도
    match hwp.register_module() {
        Ok(true) => println!("[OK] Security module registered"),
        Ok(false) => println!("[WARN] Security module registration returned false"),
        Err(e) => println!("[WARN] Security module error: {e} (continuing without it)"),
    }

    // 5. 텍스트 삽입
    hwp.insert_text("안녕하세요! COM Bridge 테스트입니다.\r\n")?;
    hwp.insert_text("Rust에서 HWP를 제어하고 있습니다.")?;
    println!("[OK] Text inserted");

    // 6. 텍스트 추출
    let text = hwp.get_text()?;
    println!("[OK] Extracted text ({} bytes, {} chars)", text.len(), text.chars().count());
    // 바이트 덤프로 실제 내용 확인
    for (i, ch) in text.chars().take(20).enumerate() {
        println!("  char[{i}] = U+{:04X} '{ch}'", ch as u32);
    }

    // 7. 저장 (보안 모듈 미등록 시 실패할 수 있음)
    let save_path = std::env::temp_dir().join("com_test_hwp_hello.hwp");
    let save_path_str = save_path.to_string_lossy().to_string();
    match hwp.save_as(&save_path_str, "HWP") {
        Ok(()) => println!("[OK] Saved to {save_path_str}"),
        Err(e) => println!("[WARN] SaveAs failed (보안 모듈 필요): {e}"),
    }

    // 8. 정리 (Drop이 Clear + Quit 호출)
    println!("[OK] Done!");

    Ok(())
}
