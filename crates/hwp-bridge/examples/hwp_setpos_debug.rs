//! SetPos/MovePos pos 파라미터 디버그
//!
//! cargo run --example hwp_setpos_debug

use com_core::{ComRuntime, DispatchObject, Variant};

fn main() -> anyhow::Result<()> {
    let _rt = ComRuntime::init()?;

    let hwp = DispatchObject::create("HWPFrame.HwpObject")?;

    // 창 표시
    let wins = hwp.get("XHwpWindows")?;
    if let Variant::Dispatch(w) = wins {
        let win0 = w.call("Item", &[Variant::I32(0)])?;
        if let Variant::Dispatch(w0) = win0 {
            w0.put("Visible", Variant::Bool(true))?;
        }
    }

    // 텍스트 삽입: ABCDEFGHIJ
    let hps = hwp.get("HParameterSet")?;
    if let Variant::Dispatch(ref hps_d) = hps {
        let hit = hps_d.get("HInsertText")?;
        if let Variant::Dispatch(ref hit_d) = hit {
            let hset = hit_d.get("HSet")?;

            let haction = hwp.get("HAction")?;
            if let Variant::Dispatch(ref ha) = haction {
                ha.call("GetDefault", &[Variant::String("InsertText".into()), hset.clone()])?;
                hit_d.put("Text", Variant::String("ABCDEFGHIJ".into()))?;
                ha.call("Execute", &[Variant::String("InsertText".into()), hit_d.get("HSet")?])?;
            }
        }
    }

    println!("텍스트: ABCDEFGHIJ\n");

    // --- SetPos 테스트 ---
    println!("=== SetPos(list, para, pos) ===");
    for pos in [0, 1, 2, 3, 5, 9] {
        let ret = hwp.call("SetPos", &[
            Variant::I32(0),   // list = 0 (본문)
            Variant::I32(0),   // para = 0 (첫 문단)
            Variant::I32(pos), // pos
        ])?;

        // MoveSelRight로 1글자 확인
        hwp.call("Run", &[Variant::String("MoveSelRight".into())])?;
        let text = hwp.call("GetTextFile", &[
            Variant::String("UNICODE".into()),
            Variant::String("saveblock".into()),
        ])?;
        hwp.call("Run", &[Variant::String("Cancel".into())])?;

        println!("  pos={pos} → ret={ret:?}, 뒤1글자={text:?}");
    }

    // --- MovePos 테스트 ---
    println!("\n=== MovePos(moveID, para, pos) ===");
    // moveID=1 (moveCurList)
    for pos in [0, 1, 2, 3, 5, 9] {
        let ret = hwp.call("MovePos", &[
            Variant::I32(1),   // moveCurList
            Variant::I32(0),   // para = 0
            Variant::I32(pos), // pos
        ])?;

        hwp.call("Run", &[Variant::String("MoveSelRight".into())])?;
        let text = hwp.call("GetTextFile", &[
            Variant::String("UNICODE".into()),
            Variant::String("saveblock".into()),
        ])?;
        hwp.call("Run", &[Variant::String("Cancel".into())])?;

        println!("  MovePos(1,0,{pos}) → ret={ret:?}, 뒤1글자={text:?}");
    }

    // --- moveID=0 (moveMain) 테스트 ---
    println!("\n=== MovePos(0, para, pos) — moveMain ===");
    for pos in [0, 3, 5, 9] {
        let ret = hwp.call("MovePos", &[
            Variant::I32(0),   // moveMain
            Variant::I32(0),   // para = 0
            Variant::I32(pos), // pos
        ])?;

        hwp.call("Run", &[Variant::String("MoveSelRight".into())])?;
        let text = hwp.call("GetTextFile", &[
            Variant::String("UNICODE".into()),
            Variant::String("saveblock".into()),
        ])?;
        hwp.call("Run", &[Variant::String("Cancel".into())])?;

        println!("  MovePos(0,0,{pos}) → ret={ret:?}, 뒤1글자={text:?}");
    }

    // --- Variant 타입 실험: pos를 다양한 타입으로 ---
    println!("\n=== SetPos — Variant 타입 실험 ===");

    // F64로 전달
    let ret = hwp.call("SetPos", &[
        Variant::I32(0),
        Variant::I32(0),
        Variant::F64(5.0),
    ])?;
    hwp.call("Run", &[Variant::String("MoveSelRight".into())])?;
    let text = hwp.call("GetTextFile", &[
        Variant::String("UNICODE".into()),
        Variant::String("saveblock".into()),
    ])?;
    hwp.call("Run", &[Variant::String("Cancel".into())])?;
    println!("  F64(5.0) → ret={ret:?}, 뒤={text:?}");

    // String으로 전달
    let ret = hwp.call("SetPos", &[
        Variant::I32(0),
        Variant::I32(0),
        Variant::String("5".into()),
    ])?;
    hwp.call("Run", &[Variant::String("MoveSelRight".into())])?;
    let text = hwp.call("GetTextFile", &[
        Variant::String("UNICODE".into()),
        Variant::String("saveblock".into()),
    ])?;
    hwp.call("Run", &[Variant::String("Cancel".into())])?;
    println!("  String(\"5\") → ret={ret:?}, 뒤={text:?}");

    // 종료
    hwp.call("Clear", &[Variant::I32(1)])?;
    hwp.call("Quit", &[])?;
    println!("\n완료");

    Ok(())
}
