use std::path::{Path, PathBuf};

use com_core::{DispatchObject, Variant};
use windows::core::*;
use windows::Win32::System::Registry::*;

// ============================================================
// HwpDetector — 한컴 설치 감지 + 보안 모듈 관리
// ============================================================

pub struct HwpDetector;

impl HwpDetector {
    /// 한컴오피스 설치 여부 확인 (COM ProgID가 등록되어 있는지)
    pub fn is_installed() -> bool {
        // CLSIDFromProgID로 확인하는 대신 레지스트리 직접 체크
        unsafe {
            let key_path = w!("HWPFrame.HwpObject\\CLSID");
            let result = RegOpenKeyExW(HKEY_CLASSES_ROOT, key_path, Some(0), KEY_READ, &mut HKEY::default());
            result.is_ok()
        }
    }

    /// 한컴오피스 설치 경로 탐색 (레지스트리 + COM 서버 경로)
    pub fn find_install_path() -> Option<PathBuf> {
        // CLSID에서 LocalServer32 경로 추출
        let clsid_paths = [
            "WOW6432Node\\CLSID\\{2291CF00-64A1-4877-A9B4-68CFE89612D6}\\LocalServer32",
            "CLSID\\{2291CF00-64A1-4877-A9B4-68CFE89612D6}\\LocalServer32",
        ];

        for path in &clsid_paths {
            if let Some(val) = read_registry_string(HKEY_CLASSES_ROOT, path, "") {
                // "C:\...\hwp.exe -Automation" → exe 경로만 추출
                let exe_path = val.trim_end_matches(" -Automation").trim_matches('"');
                let exe = Path::new(exe_path);
                if exe.exists() {
                    return exe.parent().map(|p| p.to_path_buf());
                }
            }
        }
        None
    }

    /// 보안 모듈이 레지스트리에 등록되어 있는지 확인
    pub fn is_security_module_registered() -> bool {
        let module_paths = [
            "SOFTWARE\\HNC\\HwpAutomation\\Modules",
            "SOFTWARE\\HNC\\HwpCtrl\\Modules",
        ];

        for path in &module_paths {
            if let Some(dll_path) = read_registry_string(
                HKEY_CURRENT_USER,
                path,
                "FilePathCheckerModule",
            ) {
                if Path::new(&dll_path).exists() {
                    return true;
                }
            }
        }
        false
    }

    /// 보안 모듈 DLL을 레지스트리에 등록
    /// dll_path: FilePathCheckerModuleExample.dll의 절대 경로
    pub fn register_security_module(dll_path: &str) -> Result<()> {
        let path = Path::new(dll_path);
        if !path.exists() {
            return Err(Error::new(
                HRESULT(-1),
                format!("DLL not found: {dll_path}"),
            ));
        }

        // HwpAutomation 용 등록
        write_registry_string(
            HKEY_CURRENT_USER,
            "SOFTWARE\\HNC\\HwpAutomation\\Modules",
            "FilePathCheckerModule",
            dll_path,
        )?;

        println!("[OK] Security module registered: {dll_path}");
        Ok(())
    }

    /// 진단 정보 출력
    pub fn diagnose() {
        println!("=== HWP COM 진단 ===");
        println!("설치 여부: {}", Self::is_installed());
        if let Some(path) = Self::find_install_path() {
            println!("설치 경로: {}", path.display());
        } else {
            println!("설치 경로: (찾을 수 없음)");
        }
        println!("보안 모듈: {}", if Self::is_security_module_registered() {
            "등록됨"
        } else {
            "미등록 (RegisterModule 호출 시 팝업 발생 가능)"
        });
    }
}

// ============================================================
// HwpApp — HWP COM 자동화
// ============================================================

pub struct HwpApp {
    disp: DispatchObject,
}

impl HwpApp {
    /// 새 HWP 인스턴스 실행
    pub fn new() -> Result<Self> {
        let disp = DispatchObject::create("HWPFrame.HwpObject")?;
        Ok(Self { disp })
    }

    /// Access the underlying IDispatch object
    pub fn dispatch(&self) -> &DispatchObject {
        &self.disp
    }

    /// 창 표시/숨김 (XHwpWindows.Item(0).Visible)
    pub fn set_visible(&self, visible: bool) -> Result<()> {
        let windows = self.disp.get("XHwpWindows")?.into_dispatch()?;
        let window = windows.get_by("Item", &[Variant::I32(0)])?.into_dispatch()?;
        window.put("Visible", visible)
    }

    /// 보안 모듈 등록 (COM 호출)
    /// 사전에 HwpDetector::register_security_module로 레지스트리 등록 필요
    pub fn register_module(&self) -> Result<bool> {
        let result = self.disp.call(
            "RegisterModule",
            &[
                Variant::from("FilePathCheckDLL"),
                Variant::from("FilePathCheckerModule"),
            ],
        )?;
        Ok(result.as_bool().unwrap_or(false))
    }

    /// 문서 열기
    pub fn open(&self, path: &str) -> Result<()> {
        // 절대 경로로 변환
        let abs_path = std::path::Path::new(path)
            .canonicalize()
            .unwrap_or_else(|_| PathBuf::from(path));
        let path_str = abs_path.to_string_lossy().to_string();
        // UNC prefix \\?\ 제거
        let clean_path = path_str.strip_prefix("\\\\?\\").unwrap_or(&path_str);

        self.disp.call("Open", &[Variant::from(clean_path)])?;
        Ok(())
    }

    /// 전체 텍스트 추출 (유니코드)
    pub fn get_text(&self) -> Result<String> {
        let result = self.disp.call(
            "GetTextFile",
            &[Variant::from("UNICODE"), Variant::from("")],
        )?;
        Ok(result.as_str().unwrap_or("").to_string())
    }

    /// 텍스트 삽입 (현재 커서 위치에)
    pub fn insert_text(&self, text: &str) -> Result<()> {
        // HAction.GetDefault("InsertText", HParameterSet.HInsertText.HSet)
        let haction = self.disp.get("HAction")?.into_dispatch()?;
        let hparam_set = self.disp.get("HParameterSet")?.into_dispatch()?;
        let insert_text_param = hparam_set.get("HInsertText")?.into_dispatch()?;
        let hset = insert_text_param.get("HSet")?.into_dispatch()?;

        haction.call("GetDefault", &[Variant::from("InsertText"), Variant::Dispatch(hset.clone())])?;
        insert_text_param.put("Text", text)?;
        haction.call("Execute", &[Variant::from("InsertText"), Variant::Dispatch(hset)])?;

        Ok(())
    }

    /// 다른 이름으로 저장
    /// format: "HWP", "HWPX", "PDF", "TEXT" 등
    pub fn save_as(&self, path: &str, format: &str) -> Result<()> {
        let abs_path = std::path::Path::new(path)
            .canonicalize()
            .unwrap_or_else(|_| PathBuf::from(path));
        let path_str = abs_path.to_string_lossy().to_string();
        let clean_path = path_str.strip_prefix("\\\\?\\").unwrap_or(&path_str);

        // SaveAs(fileName, format, arg) — 3개 인자 필수
        self.disp.call(
            "SaveAs",
            &[
                Variant::from(clean_path),
                Variant::from(format),
                Variant::from(""),
            ],
        )?;
        Ok(())
    }

    /// 문서 내용 비우기 (저장 안 함)
    pub fn clear(&self) -> Result<()> {
        self.disp.call("Clear", &[Variant::I32(1)])?; // 1 = 저장하지 않고 닫기
        Ok(())
    }

    /// 액션 실행 (예: "MoveDocBegin", "MoveDocEnd" 등)
    pub fn run(&self, action: &str) -> Result<()> {
        self.disp.call("Run", &[Variant::from(action)])?;
        Ok(())
    }

    /// 한글 종료
    pub fn quit(&self) -> Result<()> {
        self.disp.call("Quit", &[])?;
        Ok(())
    }
}

impl Drop for HwpApp {
    fn drop(&mut self) {
        let _ = self.clear();
        let _ = self.quit();
    }
}

// ============================================================
// Registry helpers
// ============================================================

fn read_registry_string(root: HKEY, subkey: &str, value_name: &str) -> Option<String> {
    unsafe {
        let subkey_wide: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();
        let mut hkey = HKEY::default();

        if RegOpenKeyExW(root, PCWSTR(subkey_wide.as_ptr()), Some(0), KEY_READ, &mut hkey).is_err() {
            return None;
        }

        let value_wide: Vec<u16> = value_name.encode_utf16().chain(std::iter::once(0)).collect();
        let mut data_type = REG_VALUE_TYPE::default();
        let mut data_size = 0u32;

        // 먼저 크기 조회
        if RegQueryValueExW(
            hkey,
            PCWSTR(value_wide.as_ptr()),
            None,
            Some(&mut data_type),
            None,
            Some(&mut data_size),
        ).is_err() {
            let _ = RegCloseKey(hkey);
            return None;
        }

        if data_type != REG_SZ || data_size == 0 {
            let _ = RegCloseKey(hkey);
            return None;
        }

        let mut buffer = vec![0u8; data_size as usize];
        if RegQueryValueExW(
            hkey,
            PCWSTR(value_wide.as_ptr()),
            None,
            None,
            Some(buffer.as_mut_ptr()),
            Some(&mut data_size),
        ).is_err() {
            let _ = RegCloseKey(hkey);
            return None;
        }

        let _ = RegCloseKey(hkey);

        // u8 → u16 → String
        let wide: Vec<u16> = buffer
            .chunks_exact(2)
            .map(|c| u16::from_le_bytes([c[0], c[1]]))
            .collect();
        let s = String::from_utf16_lossy(&wide);
        Some(s.trim_end_matches('\0').to_string())
    }
}

fn write_registry_string(root: HKEY, subkey: &str, value_name: &str, data: &str) -> Result<()> {
    unsafe {
        let subkey_wide: Vec<u16> = subkey.encode_utf16().chain(std::iter::once(0)).collect();
        let mut hkey = HKEY::default();

        RegCreateKeyW(
            root,
            PCWSTR(subkey_wide.as_ptr()),
            &mut hkey,
        ).ok()?;

        let value_wide: Vec<u16> = value_name.encode_utf16().chain(std::iter::once(0)).collect();
        let data_wide: Vec<u16> = data.encode_utf16().chain(std::iter::once(0)).collect();
        let data_bytes: &[u8] = std::slice::from_raw_parts(
            data_wide.as_ptr() as *const u8,
            data_wide.len() * 2,
        );

        RegSetValueExW(
            hkey,
            PCWSTR(value_wide.as_ptr()),
            Some(0),
            REG_SZ,
            Some(data_bytes),
        ).ok()?;

        let _ = RegCloseKey(hkey);
        Ok(())
    }
}
