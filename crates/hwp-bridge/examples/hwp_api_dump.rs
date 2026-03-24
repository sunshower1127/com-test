/// HWP COM 객체의 전체 API를 덤프하는 도구.
/// 루트 객체(HWPFrame.HwpObject)의 모든 메서드/프로퍼티를 열거하고,
/// Dispatch를 반환하는 프로퍼티는 한 단계 더 탐색한다.
use com_core::{ComRuntime, DispatchObject, MemberKind, Variant};

fn main() -> anyhow::Result<()> {
    let _com = ComRuntime::init()?;

    println!("=== HWP API Dump ===\n");
    println!("HWP 실행 중...");
    let hwp = DispatchObject::create("HWPFrame.HwpObject")?;

    // 창 표시
    if let Ok(windows) = hwp.get("XHwpWindows") {
        if let Ok(windows_disp) = windows.into_dispatch() {
            if let Ok(window) = windows_disp.get_by("Item", &[Variant::I32(0)]) {
                if let Ok(window_disp) = window.into_dispatch() {
                    let _ = window_disp.put("Visible", false);
                    let _ = window_disp.put("Visible", true);
                }
            }
        }
    }

    // 루트 객체 멤버 열거
    println!("\n========================================");
    println!("ROOT: HWPFrame.HwpObject");
    println!("========================================");
    let members = hwp.list_members()?;
    println!("총 {} 멤버\n", members.len());

    for m in &members {
        println!("  {m}");
    }

    // Dispatch를 반환하는 get 프로퍼티를 한 단계 더 탐색
    println!("\n\n========================================");
    println!("하위 객체 탐색 (Dispatch 반환 프로퍼티)");
    println!("========================================\n");

    let dispatch_props: Vec<_> = members
        .iter()
        .filter(|m| {
            matches!(m.kind, MemberKind::Get | MemberKind::GetPut)
                && m.return_type == "Dispatch"
                && m.params.is_empty() // 파라미터 없는 것만
        })
        .collect();

    for prop in &dispatch_props {
        println!("--- {}.{} ---", "hwp", prop.name);
        match hwp.get(&prop.name) {
            Ok(Variant::Dispatch(sub)) => {
                match sub.list_members() {
                    Ok(sub_members) => {
                        println!("  {} 멤버:", sub_members.len());
                        for sm in &sub_members {
                            println!("    {sm}");
                        }
                    }
                    Err(e) => println!("  list_members 실패: {e}"),
                }
            }
            Ok(other) => println!("  Dispatch가 아님: {other}"),
            Err(e) => println!("  get 실패: {e}"),
        }
        println!();
    }

    // CreateAction으로 주요 Action도 탐색
    println!("\n========================================");
    println!("CreateAction 탐색");
    println!("========================================\n");

    let actions = [
        "InsertText",
        "CharShape",
        "ParagraphShape",
        "TableCreate",
        "AllReplace",
        "PageSetup",
        "InsertPicture",
        "Style",
        "BreakSection",
        "InsertColumnBreak",
    ];

    for action_name in &actions {
        print!("CreateAction(\"{action_name}\"): ");
        match hwp.call("CreateAction", &[Variant::from(*action_name)]) {
            Ok(Variant::Dispatch(action_obj)) => {
                println!("OK (Dispatch)");

                // Action 객체의 멤버
                if let Ok(action_members) = action_obj.list_members() {
                    println!("  Action 멤버:");
                    for am in &action_members {
                        println!("    {am}");
                    }
                }

                // CreateSet()으로 ParameterSet 탐색
                match action_obj.call("CreateSet", &[]) {
                    Ok(Variant::Dispatch(set_obj)) => {
                        if let Ok(set_members) = set_obj.list_members() {
                            println!("  ParameterSet 멤버:");
                            for sm in &set_members {
                                println!("    {sm}");
                            }
                        }
                    }
                    Ok(other) => println!("  CreateSet 반환: {other}"),
                    Err(e) => println!("  CreateSet 실패: {e}"),
                }
            }
            Ok(other) => println!("{other}"),
            Err(e) => println!("실패 — {e}"),
        }
        println!();
    }

    // 정리
    let _ = hwp.call("Clear", &[Variant::I32(1)]);
    let _ = hwp.call("Quit", &[]);
    println!("\n=== 덤프 완료 ===");

    Ok(())
}
