mod commands;
mod state;

use com_core::ComRuntime;
use rustyline::error::ReadlineError;
use rustyline::DefaultEditor;
use state::AppState;

fn main() -> anyhow::Result<()> {
    let _com = ComRuntime::init()?;

    println!("COM CLI Explorer v0.1");
    println!("Type 'help' for commands, 'quit' to exit.\n");

    let mut rl = DefaultEditor::new()?;
    let mut state = AppState::new();

    loop {
        match rl.readline("com> ") {
            Ok(line) => {
                let _ = rl.add_history_entry(&line);
                let cmd = commands::parse(&line);
                if matches!(cmd, commands::Command::Quit) {
                    break;
                }
                if let Err(e) = commands::execute(cmd, &mut state) {
                    eprintln!("Error: {e}");
                }
            }
            Err(ReadlineError::Interrupted | ReadlineError::Eof) => break,
            Err(e) => {
                eprintln!("Error: {e}");
                break;
            }
        }
    }

    println!("Bye!");
    Ok(())
}
