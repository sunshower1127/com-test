/// 단계 0: HWP 기본 연결 — 전체 프로퍼티 + 메서드 테스트
use com_core::{ComRuntime, DispatchObject, Variant};

fn main() -> anyhow::Result<()> {
    let _com = ComRuntime::init()?;

    println!("============================================================");
    println!("  HWP 단계 0: 기본 연결 — 전체 프로퍼티/메서드 테스트");
    println!("============================================================\n");

    // ── 0-1: COM 객체 생성 ──
    println!("── 0-1: HWPFrame.HwpObject 생성 ──");
    let hwp = match DispatchObject::create("HWPFrame.HwpObject") {
        Ok(h) => { println!("  ✅ 성공"); h }
        Err(e) => { println!("  ❌ 실패: {e}"); return Ok(()); }
    };

    // ── 0-2: 창 표시 ──
    println!("\n── 0-2: XHwpWindows.Item(0).Visible = true ──");
    match show_window(&hwp) {
        Ok(()) => println!("  ✅ 성공"),
        Err(e) => println!("  ❌ 실패: {e}"),
    }

    // ══════════════════════════════════════════════
    //  단순 스칼라 프로퍼티 (get)
    // ══════════════════════════════════════════════
    println!("\n══ 스칼라 프로퍼티 (get) ══\n");

    test_get(&hwp, "0-3", "Version", "String — 버전 문자열");
    test_get(&hwp, "0-4", "EditMode", "I4 — 0=읽기전용, 1=일반, 2=양식, 16=배포");
    test_get(&hwp, "0-5", "IsEmpty", "Bool — 빈 문서 여부");
    test_get(&hwp, "0-6", "PageCount", "I4 — 페이지 수");
    test_get(&hwp, "0-7", "IsModified", "Bool — 수정 여부");
    test_get(&hwp, "0-8", "IsPrivateInfoProtected", "Bool — 개인정보 보호");
    test_get(&hwp, "0-9", "Path", "String — 현재 문서 경로");
    test_get(&hwp, "0-10", "SelectionMode", "I4 — 선택 모드");
    test_get(&hwp, "0-11", "CurFieldState", "I4 — 현재 필드 상태");

    // ══════════════════════════════════════════════
    //  공식 문서에 있지만 ITypeInfo 덤프에 없던 프로퍼티
    // ══════════════════════════════════════════════
    println!("\n══ 공식 문서에만 있는 프로퍼티 (덤프에 누락?) ══\n");

    test_get(&hwp, "0-12", "CLSID", "공식 문서에 존재 — 덤프에 없음");
    test_get(&hwp, "0-13", "CurMetatagState", "공식 문서에 존재 — 덤프에 없음");
    test_get(&hwp, "0-14", "IsTrackChange", "공식 문서에 존재 — 덤프에 없음");
    test_get(&hwp, "0-15", "IsTrackChangePassword", "공식 문서에 존재 — 덤프에 없음");

    // ══════════════════════════════════════════════
    //  Dispatch 반환 프로퍼티 — 접근 가능 여부
    // ══════════════════════════════════════════════
    println!("\n══ Dispatch 프로퍼티 — 접근 가능 여부 ══\n");

    let dispatch_props = [
        ("0-16", "Application", "자기 참조 (루트와 동일?)"),
        ("0-17", "CharShape", "현재 글자 모양 ParameterSet"),
        ("0-18", "ParaShape", "현재 문단 모양 ParameterSet"),
        ("0-19", "CellShape", "현재 셀 모양 (빈 문서=empty?)"),
        ("0-20", "HAction", "HAction 객체"),
        ("0-21", "HParameterSet", "HParameterSet (429 멤버)"),
        ("0-22", "EngineProperties", "엔진 속성 ParameterSet"),
        ("0-23", "ViewProperties", "보기 속성 ParameterSet"),
        ("0-24", "HeadCtrl", "문서 첫 컨트롤 (빈 문서=empty?)"),
        ("0-25", "LastCtrl", "문서 마지막 컨트롤 (빈 문서=empty?)"),
        ("0-26", "CurSelectedCtrl", "선택된 컨트롤 (없으면 empty?)"),
        ("0-27", "ParentCtrl", "부모 컨트롤 (없으면 empty?)"),
        ("0-28", "XHwpDocuments", "문서 컬렉션"),
        ("0-29", "XHwpWindows", "창 컬렉션"),
        ("0-30", "XHwpMessageBox", "메시지박스 객체"),
        ("0-31", "XHwpODBC", "ODBC 객체"),
    ];

    for (id, name, desc) in &dispatch_props {
        test_get_dispatch(&hwp, id, name, desc);
    }

    // ══════════════════════════════════════════════
    //  Dispatch 프로퍼티 상세 탐색 — 컬렉션 카운트 등
    // ══════════════════════════════════════════════
    println!("\n══ 컬렉션 상세 탐색 ══\n");

    // XHwpDocuments.Count
    println!("── 0-32: XHwpDocuments.Count ──");
    match (|| -> anyhow::Result<()> {
        let docs = hwp.get("XHwpDocuments")?.into_dispatch()?;
        let count = docs.get("Count")?;
        println!("  ✅ Count = {count}");
        Ok(())
    })() {
        Ok(()) => {}
        Err(e) => println!("  ❌ {e}"),
    }

    // XHwpWindows.Count
    println!("── 0-33: XHwpWindows.Count ──");
    match (|| -> anyhow::Result<()> {
        let wins = hwp.get("XHwpWindows")?.into_dispatch()?;
        let count = wins.get("Count")?;
        println!("  ✅ Count = {count}");
        Ok(())
    })() {
        Ok(()) => {}
        Err(e) => println!("  ❌ {e}"),
    }

    // XHwpDocuments.Active_XHwpDocument
    println!("── 0-34: XHwpDocuments.Active_XHwpDocument ──");
    match (|| -> anyhow::Result<()> {
        let docs = hwp.get("XHwpDocuments")?.into_dispatch()?;
        let active = docs.get("Active_XHwpDocument")?;
        match active {
            Variant::Dispatch(_) => println!("  ✅ Dispatch 반환"),
            other => println!("  ⚠️ 반환: {other}"),
        }
        Ok(())
    })() {
        Ok(()) => {}
        Err(e) => println!("  ❌ {e}"),
    }

    // XHwpWindows.Active_XHwpWindow
    println!("── 0-35: XHwpWindows.Active_XHwpWindow ──");
    match (|| -> anyhow::Result<()> {
        let wins = hwp.get("XHwpWindows")?.into_dispatch()?;
        let active = wins.get("Active_XHwpWindow")?;
        match active {
            Variant::Dispatch(_) => println!("  ✅ Dispatch 반환"),
            other => println!("  ⚠️ 반환: {other}"),
        }
        Ok(())
    })() {
        Ok(()) => {}
        Err(e) => println!("  ❌ {e}"),
    }

    // ══════════════════════════════════════════════
    //  ParameterSet 프로퍼티 상세 — Item 접근
    // ══════════════════════════════════════════════
    println!("\n══ ParameterSet 프로퍼티 상세 탐색 ══\n");

    // CharShape — SetID 확인 + Item 접근
    println!("── 0-36: CharShape.SetID + Item(\"Height\") ──");
    match (|| -> anyhow::Result<()> {
        let cs = hwp.get("CharShape")?.into_dispatch()?;
        let set_id = cs.get("SetID")?;
        println!("  SetID = {set_id}");
        let height = cs.call("Item", &[Variant::from("Height")])?;
        println!("  ✅ Item(\"Height\") = {height}");
        let bold = cs.call("Item", &[Variant::from("Bold")])?;
        println!("  ✅ Item(\"Bold\") = {bold}");
        Ok(())
    })() {
        Ok(()) => {}
        Err(e) => println!("  ❌ {e}"),
    }

    // ParaShape — SetID + Item
    println!("── 0-37: ParaShape.SetID + Item(\"Alignment\") ──");
    match (|| -> anyhow::Result<()> {
        let ps = hwp.get("ParaShape")?.into_dispatch()?;
        let set_id = ps.get("SetID")?;
        println!("  SetID = {set_id}");
        let align = ps.call("Item", &[Variant::from("Alignment")])?;
        println!("  ✅ Item(\"Alignment\") = {align}");
        Ok(())
    })() {
        Ok(()) => {}
        Err(e) => println!("  ❌ {e}"),
    }

    // HParameterSet — 하위 접근 (HCharShape)
    println!("── 0-38: HParameterSet.HCharShape 접근 ──");
    match (|| -> anyhow::Result<()> {
        let hps = hwp.get("HParameterSet")?.into_dispatch()?;
        let hcs = hps.get("HCharShape")?;
        match &hcs {
            Variant::Dispatch(d) => {
                let set_id = d.get("SetID")?;
                println!("  ✅ Dispatch 반환, SetID = {set_id}");
            }
            other => println!("  ⚠️ 반환: {other}"),
        }
        Ok(())
    })() {
        Ok(()) => {}
        Err(e) => println!("  ❌ {e} — HParameterSet 직접 접근 실패 여부 확인!"),
    }

    // HParameterSet — HInsertText
    println!("── 0-39: HParameterSet.HInsertText 접근 ──");
    match (|| -> anyhow::Result<()> {
        let hps = hwp.get("HParameterSet")?.into_dispatch()?;
        let hit = hps.get("HInsertText")?;
        match &hit {
            Variant::Dispatch(d) => {
                let set_id = d.get("SetID")?;
                println!("  ✅ Dispatch 반환, SetID = {set_id}");
            }
            other => println!("  ⚠️ 반환: {other}"),
        }
        Ok(())
    })() {
        Ok(()) => {}
        Err(e) => println!("  ❌ {e}"),
    }

    // HParameterSet — HTableCreation
    println!("── 0-40: HParameterSet.HTableCreation 접근 ──");
    match (|| -> anyhow::Result<()> {
        let hps = hwp.get("HParameterSet")?.into_dispatch()?;
        let htc = hps.get("HTableCreation")?;
        match &htc {
            Variant::Dispatch(d) => {
                let set_id = d.get("SetID")?;
                println!("  ✅ Dispatch 반환, SetID = {set_id}");
            }
            other => println!("  ⚠️ 반환: {other}"),
        }
        Ok(())
    })() {
        Ok(()) => {}
        Err(e) => println!("  ❌ {e}"),
    }

    // ══════════════════════════════════════════════
    //  EditMode 쓰기 테스트
    // ══════════════════════════════════════════════
    println!("\n══ EditMode 쓰기 테스트 ══\n");

    println!("── 0-41: EditMode = 0 (읽기전용) → 복원 ──");
    match (|| -> anyhow::Result<()> {
        let original = hwp.get("EditMode")?;
        println!("  현재 EditMode = {original}");
        hwp.put("EditMode", 0)?;
        let after = hwp.get("EditMode")?;
        println!("  put(0) 후 EditMode = {after}");
        // 복원
        hwp.put("EditMode", 1)?;
        let restored = hwp.get("EditMode")?;
        println!("  ✅ 복원 후 EditMode = {restored}");
        Ok(())
    })() {
        Ok(()) => {}
        Err(e) => println!("  ❌ {e}"),
    }

    // ══════════════════════════════════════════════
    //  SetMessageBoxMode / GetMessageBoxMode
    // ══════════════════════════════════════════════
    println!("\n══ SetMessageBoxMode 테스트 ══\n");

    let modes: &[(&str, i32)] = &[
        ("0x00010000 (자동 예)", 0x00010000),
        ("0x00001000 (예 클릭)", 0x00001000),
        ("0x00002000 (아니오 클릭)", 0x00002000),
        ("0x00004000 (취소 클릭)", 0x00004000),
    ];

    for (desc, mode) in modes {
        print!("── Set({desc}) → ");
        match hwp.call("SetMessageBoxMode", &[Variant::I32(*mode)]) {
            Ok(ret) => {
                let current = hwp.call("GetMessageBoxMode", &[]).unwrap_or(Variant::Empty);
                println!("✅ 반환={ret}, Get={current}");
            }
            Err(e) => println!("❌ {e}"),
        }
    }

    // 리셋
    println!("── 리셋(0x000F0000) → ");
    match hwp.call("SetMessageBoxMode", &[Variant::I32(0x000F0000)]) {
        Ok(ret) => {
            let current = hwp.call("GetMessageBoxMode", &[]).unwrap_or(Variant::Empty);
            println!("  ✅ 반환={ret}, Get={current}");
        }
        Err(e) => println!("  ❌ {e}"),
    }

    // ══════════════════════════════════════════════
    //  유틸리티 메서드 (독립 실행 가능)
    // ══════════════════════════════════════════════
    println!("\n══ 유틸리티 메서드 ══\n");

    println!("── 0-42: RGBColor(255, 0, 0) ──");
    match hwp.call("RGBColor", &[Variant::I32(255), Variant::I32(0), Variant::I32(0)]) {
        Ok(v) => println!("  ✅ = {v}"),
        Err(e) => println!("  ❌ {e}"),
    }

    println!("── 0-43: RGBColor(0, 0, 255) ──");
    match hwp.call("RGBColor", &[Variant::I32(0), Variant::I32(0), Variant::I32(255)]) {
        Ok(v) => println!("  ✅ = {v}"),
        Err(e) => println!("  ❌ {e}"),
    }

    println!("── 0-44: MiliToHwpUnit(25.4) — 1인치 ──");
    match hwp.call("MiliToHwpUnit", &[Variant::F64(25.4)]) {
        Ok(v) => println!("  ✅ = {v}"),
        Err(e) => println!("  ❌ {e}"),
    }

    println!("── 0-45: PointToHwpUnit(12.0) — 12pt ──");
    match hwp.call("PointToHwpUnit", &[Variant::F64(12.0)]) {
        Ok(v) => println!("  ✅ = {v}"),
        Err(e) => println!("  ❌ {e}"),
    }

    println!("── 0-46: HwpLineType(\"Solid\") ──");
    match hwp.call("HwpLineType", &[Variant::from("Solid")]) {
        Ok(v) => println!("  ✅ = {v}"),
        Err(e) => println!("  ❌ {e}"),
    }

    println!("── 0-47: HAlign(\"Center\") ──");
    match hwp.call("HAlign", &[Variant::from("Center")]) {
        Ok(v) => println!("  ✅ = {v}"),
        Err(e) => println!("  ❌ {e}"),
    }

    println!("── 0-48: TextAlign(\"Center\") ──");
    match hwp.call("TextAlign", &[Variant::from("Center")]) {
        Ok(v) => println!("  ✅ = {v}"),
        Err(e) => println!("  ❌ {e}"),
    }

    // ══════════════════════════════════════════════
    //  IsActionEnable — 주요 액션 사용 가능 여부
    // ══════════════════════════════════════════════
    println!("\n══ IsActionEnable — 주요 액션 확인 ══\n");

    let actions = [
        "InsertText", "CharShape", "ParagraphShape", "TableCreate",
        "AllReplace", "PageSetup", "InsertPicture", "Style",
        "BreakSection", "InsertColumnBreak", "Delete", "Copy",
        "Cut", "Paste", "Undo", "Redo", "FileNew", "FileOpen",
        "FileSaveAs", "Print",
    ];

    for act in &actions {
        match hwp.call("IsActionEnable", &[Variant::from(*act)]) {
            Ok(v) => println!("  {act:20} → {v}"),
            Err(e) => println!("  {act:20} → ❌ {e}"),
        }
    }

    // ══════════════════════════════════════════════
    //  완료
    // ══════════════════════════════════════════════
    println!("\n============================================================");
    println!("  단계 0 전체 테스트 완료!");
    println!("  한글 창이 보이는지 확인해주세요.");
    println!("  Enter 키를 누르면 한글을 종료합니다...");
    println!("============================================================");
    let mut buf = String::new();
    std::io::stdin().read_line(&mut buf).ok();

    let _ = hwp.call("Clear", &[Variant::I32(1)]);
    let _ = hwp.call("Quit", &[]);

    Ok(())
}

fn show_window(hwp: &DispatchObject) -> anyhow::Result<()> {
    let windows = hwp.get("XHwpWindows")?.into_dispatch()?;
    let window = windows.get_by("Item", &[Variant::I32(0)])?.into_dispatch()?;
    let _ = window.put("Visible", false);
    window.put("Visible", true)?;
    Ok(())
}

fn test_get(hwp: &DispatchObject, id: &str, name: &str, desc: &str) {
    print!("── {id}: {name} — {desc} ── ");
    match hwp.get(name) {
        Ok(v) => println!("✅ = {v}"),
        Err(e) => println!("❌ {e}"),
    }
}

fn test_get_dispatch(hwp: &DispatchObject, id: &str, name: &str, desc: &str) {
    print!("── {id}: {name} — {desc} ── ");
    match hwp.get(name) {
        Ok(Variant::Dispatch(d)) => {
            // 멤버 수 확인
            match d.list_members() {
                Ok(members) => println!("✅ Dispatch ({} 멤버)", members.len()),
                Err(_) => println!("✅ Dispatch (멤버 열거 실패)"),
            }
        }
        Ok(Variant::Empty) => println!("⚠️ empty (빈 문서라서 가능)"),
        Ok(other) => println!("⚠️ Dispatch 아님: {other}"),
        Err(e) => println!("❌ {e}"),
    }
}
