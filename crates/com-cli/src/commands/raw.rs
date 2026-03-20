use anyhow::{bail, Result};
use com_core::{DispatchObject, Variant};

use crate::state::AppState;

pub fn handle(args: &[String], state: &mut AppState) -> Result<()> {
    let action = args.first().map(|s| s.as_str()).unwrap_or("help");

    match action {
        "target" => {
            let target_name = args.get(1).map(|s| s.as_str()).unwrap_or("");
            match target_name {
                "excel" | "xl" => {
                    let app = state.excel.as_ref()
                        .ok_or_else(|| anyhow::anyhow!("Excel is not running."))?;
                    state.raw_target = Some(app.dispatch().clone());
                    println!("[OK] Raw target: Excel app");
                }
                "hwp" => {
                    let app = state.hwp.as_ref()
                        .ok_or_else(|| anyhow::anyhow!("HWP is not running."))?;
                    state.raw_target = Some(app.dispatch().clone());
                    println!("[OK] Raw target: HWP app");
                }
                "result" => {
                    if let Some(ref d) = state.last_result {
                        state.raw_target = Some(d.clone());
                        println!("[OK] Raw target: last result");
                    } else {
                        bail!("No last result to target.");
                    }
                }
                "workbook" | "wb" => {
                    let wb = state.excel_workbook.as_ref()
                        .ok_or_else(|| anyhow::anyhow!("No workbook open."))?;
                    state.raw_target = Some(wb.dispatch().clone());
                    println!("[OK] Raw target: Excel workbook");
                }
                "sheet" => {
                    let sh = state.excel_sheet.as_ref()
                        .ok_or_else(|| anyhow::anyhow!("No sheet selected."))?;
                    state.raw_target = Some(sh.dispatch().clone());
                    println!("[OK] Raw target: Excel sheet");
                }
                _ => bail!("Usage: raw target excel|hwp|result|workbook|sheet"),
            }
        }

        "get" => {
            let prop = require_arg(args, 1, "raw get <Property>")?;
            let target = require_target(state)?;
            let result = target.get(&prop)?;
            print_result(&result, state);
        }

        "getby" => {
            let prop = require_arg(args, 1, "raw getby <Property> <args...>")?;
            let target = require_target(state)?;
            let variant_args: Vec<Variant> = args[2..].iter().map(|s| parse_value(s)).collect();
            let result = target.get_by(&prop, &variant_args)?;
            print_result(&result, state);
        }

        "put" => {
            let prop = require_arg(args, 1, "raw put <Property> <value>")?;
            let val_str = require_arg(args, 2, "raw put <Property> <value>")?;
            let target = require_target(state)?;
            let value = parse_value(&val_str);
            target.put(&prop, value)?;
            println!("[OK] {prop} set.");
        }

        "call" => {
            let method = require_arg(args, 1, "raw call <Method> [args...]")?;
            let target = require_target(state)?;
            let variant_args: Vec<Variant> = args[2..].iter().map(|s| parse_value(s)).collect();
            let result = target.call(&method, &variant_args)?;
            print_result(&result, state);
        }

        "chain" => {
            let chain_str = require_arg(args, 1, "raw chain <Prop1.Prop2.Prop3>")?;
            let parts: Vec<&str> = chain_str.split('.').collect();
            let mut current = require_target(state)?.clone();

            for (i, prop) in parts.iter().enumerate() {
                let result = current.get(prop)?;
                if i == parts.len() - 1 {
                    // 마지막 프로퍼티: 결과 출력
                    print_result(&result, state);
                } else {
                    // 중간 프로퍼티: dispatch로 이동
                    current = result.into_dispatch()
                        .map_err(|_| anyhow::anyhow!("'{prop}' is not a dispatch object, cannot continue chain"))?;
                    println!("  .{prop} => (dispatch)");
                }
            }
        }

        _ => println!("Unknown raw command: {action} (type 'help')"),
    }

    Ok(())
}

fn require_target(state: &AppState) -> Result<&DispatchObject> {
    state.raw_target.as_ref()
        .ok_or_else(|| anyhow::anyhow!("No raw target set. Use: raw target excel|hwp"))
}

fn require_arg(args: &[String], index: usize, usage: &str) -> Result<String> {
    args.get(index)
        .cloned()
        .ok_or_else(|| anyhow::anyhow!("Usage: {usage}"))
}

fn print_result(result: &Variant, state: &mut AppState) {
    println!("  => {result}");
    if let Variant::Dispatch(d) = result {
        state.last_result = Some(d.clone());
        println!("  (dispatch object stored — use 'raw target result' to explore)");
    }
}

fn parse_value(s: &str) -> Variant {
    if let Ok(n) = s.parse::<i32>() {
        return Variant::I32(n);
    }
    if let Ok(n) = s.parse::<f64>() {
        return Variant::F64(n);
    }
    match s {
        "true" => Variant::Bool(true),
        "false" => Variant::Bool(false),
        _ => Variant::String(s.to_string()),
    }
}
