/// HParameterSet 하위 객체 내부 구조 조사
/// list_members()로 실제 어떤 멤버가 있는지 확인
use com_core::{ComRuntime, DispatchObject, Variant};

fn main() -> anyhow::Result<()> {
    let _com = ComRuntime::init()?;
    let hwp = DispatchObject::create("HWPFrame.HwpObject")?;

    let hps = hwp.get("HParameterSet")?.into_dispatch()?;

    // 1) HParameterSet 자체의 list_members
    println!("=== HParameterSet 자체 (list_members) ===");
    match hps.list_members() {
        Ok(members) => {
            println!("  {} 멤버 (처음 20개만 출력)", members.len());
            for m in members.iter().take(20) {
                println!("    {m}");
            }
        }
        Err(e) => println!("  실패: {e}"),
    }

    // 2) HCharShape 가져오기 + list_members
    println!("\n=== HParameterSet.HCharShape → list_members ===");
    match hps.get("HCharShape") {
        Ok(Variant::Dispatch(d)) => {
            println!("  get 성공 — Dispatch 반환");
            match d.list_members() {
                Ok(members) => {
                    println!("  {} 멤버:", members.len());
                    for m in &members {
                        println!("    {m}");
                    }
                }
                Err(e) => println!("  list_members 실패: {e}"),
            }
            // 직접 프로퍼티 이름들 시도
            println!("\n  직접 접근 시도:");
            for name in &["SetID", "Item", "Height", "Bold", "TextColor", "FontName", "Count", "IsSet"] {
                match d.get(name) {
                    Ok(v) => println!("    get(\"{name}\") = {v}"),
                    Err(_) => {
                        match d.call(name, &[]) {
                            Ok(v) => println!("    call(\"{name}\") = {v}"),
                            Err(e2) => println!("    {name} → ❌ {e2}"),
                        }
                    }
                }
            }
        }
        Ok(other) => println!("  get 반환: {other}"),
        Err(e) => println!("  get 실패: {e}"),
    }

    // 3) get_by_typeinfo로 가져온 것도 비교
    println!("\n=== get_by_typeinfo(\"HCharShape\") → list_members ===");
    match hps.get_by_typeinfo("HCharShape") {
        Ok(Variant::Dispatch(d)) => {
            println!("  get_by_typeinfo 성공 — Dispatch 반환");
            match d.list_members() {
                Ok(members) => {
                    println!("  {} 멤버:", members.len());
                    for m in &members {
                        println!("    {m}");
                    }
                }
                Err(e) => println!("  list_members 실패: {e}"),
            }
        }
        Ok(other) => println!("  반환: {other}"),
        Err(e) => println!("  실패: {e}"),
    }

    // 4) HInsertText도 확인
    println!("\n=== HParameterSet.HInsertText → list_members ===");
    match hps.get("HInsertText") {
        Ok(Variant::Dispatch(d)) => {
            match d.list_members() {
                Ok(members) => {
                    println!("  {} 멤버:", members.len());
                    for m in &members {
                        println!("    {m}");
                    }
                }
                Err(e) => println!("  list_members 실패: {e}"),
            }
            // Text 프로퍼티 직접 시도
            for name in &["Text", "SetID", "HSet"] {
                match d.get(name) {
                    Ok(v) => println!("  get(\"{name}\") = {v}"),
                    Err(e) => println!("  get(\"{name}\") → ❌ {e}"),
                }
            }
        }
        Ok(other) => println!("  반환: {other}"),
        Err(e) => println!("  실패: {e}"),
    }

    // 5) CreateAction으로 만든 Set과 비교
    println!("\n=== CreateAction(\"InsertText\").CreateSet() → list_members ===");
    match hwp.call("CreateAction", &[Variant::from("InsertText")]) {
        Ok(Variant::Dispatch(act)) => {
            match act.call("CreateSet", &[]) {
                Ok(Variant::Dispatch(set)) => {
                    match set.list_members() {
                        Ok(members) => {
                            println!("  {} 멤버:", members.len());
                            for m in &members {
                                println!("    {m}");
                            }
                        }
                        Err(e) => println!("  list_members 실패: {e}"),
                    }
                    for name in &["SetID", "Count", "IsSet"] {
                        match set.get(name) {
                            Ok(v) => println!("  get(\"{name}\") = {v}"),
                            Err(e) => println!("  get(\"{name}\") → ❌ {e}"),
                        }
                    }
                }
                Ok(other) => println!("  CreateSet 반환: {other}"),
                Err(e) => println!("  CreateSet 실패: {e}"),
            }
        }
        Ok(_) => println!("  CreateAction 반환이 Dispatch 아님"),
        Err(e) => println!("  CreateAction 실패: {e}"),
    }

    let _ = hwp.call("Clear", &[Variant::I32(1)]);
    let _ = hwp.call("Quit", &[]);
    Ok(())
}
