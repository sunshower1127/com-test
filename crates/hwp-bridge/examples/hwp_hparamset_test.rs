/// HParameterSet 접근 실험:
/// GetIDsOfNames (late-binding) vs ITypeInfo DISPID 직접 사용 비교
use com_core::{ComRuntime, DispatchObject, Variant};

fn main() -> anyhow::Result<()> {
    let _com = ComRuntime::init()?;

    println!("=== HParameterSet 접근 실험 ===\n");

    let hwp = DispatchObject::create("HWPFrame.HwpObject")?;
    println!("HWP 실행 완료\n");

    // 1. HParameterSet 자체를 가져오기
    println!("── 1: hwp.get(\"HParameterSet\") ──");
    let hps = match hwp.get("HParameterSet") {
        Ok(Variant::Dispatch(d)) => {
            println!("  ✅ Dispatch 반환");
            d
        }
        Ok(other) => {
            println!("  ❌ Dispatch가 아님: {other}");
            return Ok(());
        }
        Err(e) => {
            println!("  ❌ {e}");
            return Ok(());
        }
    };

    // 2. 방법 A: GetIDsOfNames (일반 late-binding) — 실패 예상
    println!("\n── 2A: hps.get(\"HCharShape\") — GetIDsOfNames 경유 ──");
    match hps.get("HCharShape") {
        Ok(v) => println!("  ✅ 성공! = {v}"),
        Err(e) => println!("  ❌ 실패 (예상대로): {e}"),
    }

    // 3. 방법 B: ITypeInfo에서 DISPID 직접 → Invoke
    println!("\n── 2B: hps.get_by_typeinfo(\"HCharShape\") — DISPID 직접 사용 ──");
    match hps.get_by_typeinfo("HCharShape") {
        Ok(Variant::Dispatch(d)) => {
            println!("  ✅ Dispatch 반환!");
            // SetID 확인
            match d.get("SetID") {
                Ok(v) => println!("  SetID = {v}"),
                Err(e) => println!("  SetID 실패: {e}"),
            }
            // Item 접근
            match d.call("Item", &[Variant::from("Height")]) {
                Ok(v) => println!("  Item(\"Height\") = {v}"),
                Err(e) => println!("  Item(\"Height\") 실패: {e}"),
            }
        }
        Ok(other) => println!("  ⚠️ Dispatch 아님: {other}"),
        Err(e) => println!("  ❌ 실패: {e}"),
    }

    // 4. 다른 HParameterSet 하위 프로퍼티도 시도
    let sub_props = ["HInsertText", "HTableCreation", "HParaShape", "HFindReplace", "HPageDef"];
    for name in &sub_props {
        println!("\n── hps.get_by_typeinfo(\"{name}\") ──");
        match hps.get_by_typeinfo(name) {
            Ok(Variant::Dispatch(d)) => {
                match d.get("SetID") {
                    Ok(v) => println!("  ✅ Dispatch, SetID = {v}"),
                    Err(_) => println!("  ✅ Dispatch (SetID 읽기 실패)"),
                }
            }
            Ok(other) => println!("  ⚠️ {other}"),
            Err(e) => println!("  ❌ {e}"),
        }
    }

    // 5. 만약 성공했다면: HParameterSet 패턴으로 InsertText 시도
    println!("\n── 5: HParameterSet 패턴으로 InsertText 시도 ──");
    match (|| -> anyhow::Result<()> {
        let hit = hps.get_by_typeinfo("HInsertText")?.into_dispatch()?;
        println!("  HInsertText 획득, SetID = {}", hit.get("SetID")?);

        // HSet 가져오기
        let hset = hit.get_by_typeinfo("HSet");
        println!("  HSet via typeinfo: {:?}", hset.as_ref().map(|v| format!("{v}")).unwrap_or_else(|e| format!("err: {e}")));

        // 일반 get도 시도
        let hset2 = hit.get("HSet");
        println!("  HSet via get: {:?}", hset2.as_ref().map(|v| format!("{v}")).unwrap_or_else(|e| format!("err: {e}")));

        Ok(())
    })() {
        Ok(()) => println!("  완료"),
        Err(e) => println!("  ❌ {e}"),
    }

    println!("\nEnter 키를 누르면 종료합니다...");
    let mut buf = String::new();
    std::io::stdin().read_line(&mut buf).ok();

    let _ = hwp.call("Clear", &[Variant::I32(1)]);
    let _ = hwp.call("Quit", &[]);
    Ok(())
}
