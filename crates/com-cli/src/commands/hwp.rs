use anyhow::{bail, Result};
use hwp_bridge::{HwpApp, HwpDetector};

use crate::state::AppState;

pub fn handle(args: &[String], state: &mut AppState) -> Result<()> {
    let action = args.first().map(|s| s.as_str()).unwrap_or("help");

    match action {
        "detect" => {
            HwpDetector::diagnose();
        }

        "launch" => {
            if state.hwp.is_some() {
                println!("HWP is already running.");
                return Ok(());
            }
            let app = HwpApp::new()?;
            app.set_visible(true)?;
            state.hwp = Some(app);
            println!("[OK] HWP launched.");
        }

        "security" => {
            let dll_path = require_arg(args, 1, "hwp security <dll_path>")?;
            HwpDetector::register_security_module(&dll_path)?;
            if let Some(ref hwp) = state.hwp {
                let ok = hwp.register_module()?;
                println!("[OK] RegisterModule result: {ok}");
            } else {
                println!("[OK] DLL registered. Launch HWP to call RegisterModule.");
            }
        }

        "open" => {
            let path = require_arg(args, 1, "hwp open <path>")?;
            let hwp = require_hwp(state)?;
            hwp.open(&path)?;
            state.hwp_file = Some(path.clone());
            println!("[OK] Opened: {path}");
        }

        "text" => {
            let hwp = require_hwp(state)?;
            let text = hwp.get_text()?;
            println!("--- text ---");
            println!("{text}");
            println!("--- end ---");
        }

        "insert" => {
            let text = args.get(1..).map(|s| s.join(" ")).unwrap_or_default();
            if text.is_empty() {
                bail!("Usage: hwp insert <text>");
            }
            let hwp = require_hwp(state)?;
            hwp.insert_text(&text)?;
            println!("[OK] Inserted.");
        }

        "run" => {
            let action_name = require_arg(args, 1, "hwp run <action>")?;
            let hwp = require_hwp(state)?;
            hwp.run(&action_name)?;
            println!("[OK] Run: {action_name}");
        }

        "save" => {
            let path = require_arg(args, 1, "hwp save <path> [format]")?;
            let format = args.get(2).map(|s| s.as_str()).unwrap_or("HWP");
            let hwp = require_hwp(state)?;
            hwp.save_as(&path, format)?;
            println!("[OK] Saved to {path} (format: {format})");
        }

        "clear" => {
            let hwp = require_hwp(state)?;
            hwp.clear()?;
            state.hwp_file = None;
            println!("[OK] Document cleared.");
        }

        "quit" => {
            if let Some(hwp) = state.hwp.take() {
                let _ = hwp.clear();
                let _ = hwp.quit();
                std::mem::forget(hwp); // Drop에서 중복 호출 방지
            }
            state.clear_hwp();
            println!("[OK] HWP quit.");
        }

        _ => println!("Unknown hwp command: {action} (type 'help')"),
    }

    Ok(())
}

fn require_hwp(state: &AppState) -> Result<&HwpApp> {
    state.hwp.as_ref().ok_or_else(|| anyhow::anyhow!("HWP is not running. Use 'hwp launch' first."))
}

fn require_arg(args: &[String], index: usize, usage: &str) -> Result<String> {
    args.get(index)
        .cloned()
        .ok_or_else(|| anyhow::anyhow!("Usage: {usage}"))
}
