pub mod excel;
pub mod hwp;
pub mod raw;

use crate::state::AppState;
use anyhow::Result;

pub enum Command {
    Help,
    Status,
    Quit,
    Excel(Vec<String>),
    Hwp(Vec<String>),
    Raw(Vec<String>),
    Unknown(String),
    Empty,
}

pub fn parse(input: &str) -> Command {
    let trimmed = input.trim();
    if trimmed.is_empty() {
        return Command::Empty;
    }

    let mut parts = trimmed.splitn(2, char::is_whitespace);
    let namespace = parts.next().unwrap_or("");
    let rest: Vec<String> = match parts.next() {
        Some(s) => s.split_whitespace().map(String::from).collect(),
        None => vec![],
    };

    match namespace {
        "help" | "h" | "?" => Command::Help,
        "status" | "st" => Command::Status,
        "quit" | "exit" | "q" => Command::Quit,
        "excel" | "xl" => Command::Excel(rest),
        "hwp" => Command::Hwp(rest),
        "raw" | "r" => Command::Raw(rest),
        _ => Command::Unknown(trimmed.to_string()),
    }
}

pub fn execute(cmd: Command, state: &mut AppState) -> Result<()> {
    match cmd {
        Command::Help => print_help(),
        Command::Status => state.print_status(),
        Command::Excel(args) => excel::handle(&args, state)?,
        Command::Hwp(args) => hwp::handle(&args, state)?,
        Command::Raw(args) => raw::handle(&args, state)?,
        Command::Unknown(s) => println!("Unknown command: {s} (type 'help')"),
        Command::Empty | Command::Quit => {}
    }
    Ok(())
}

fn print_help() {
    println!(
        r#"
=== COM CLI Explorer ===

  help / h / ?          이 도움말
  status / st           현재 상태 확인
  quit / exit / q       종료

--- Excel (excel / xl) ---
  excel launch          Excel 실행
  excel new             새 워크북
  excel open <path>     파일 열기
  excel sheet [n]       시트 전환 (기본 1)
  excel cell <A1>       셀 값 읽기
  excel set <A1> <val>  셀 값 쓰기
  excel formula <A1> <expr>   수식 설정
  excel getformula <A1> 수식 읽기
  excel save <path>     저장
  excel close           워크북 닫기
  excel quit            Excel 종료

--- HWP (hwp) ---
  hwp detect            설치/보안모듈 상태 확인
  hwp launch            HWP 실행
  hwp security <dll>    보안 모듈 등록
  hwp open <path>       문서 열기
  hwp text              전체 텍스트 추출
  hwp insert <text>     텍스트 삽입
  hwp run <action>      HWP 액션 (MoveDocBegin 등)
  hwp save <path> [fmt] 저장 (HWP/HWPX/PDF/TEXT)
  hwp clear             문서 닫기 (저장 안 함)
  hwp quit              HWP 종료

--- Raw dispatch (raw / r) ---
  raw target excel|hwp|result  대상 설정
  raw get <Property>    프로퍼티 읽기
  raw getby <Prop> <args>  인덱스 프로퍼티
  raw put <Prop> <val>  프로퍼티 설정
  raw call <Method> [args]  메서드 호출
  raw chain <A.B.C>     프로퍼티 체인 탐색
"#
    );
}
