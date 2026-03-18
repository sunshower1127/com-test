use std::fmt;
use std::mem::ManuallyDrop;

use windows::core::*;
use windows::Win32::Foundation::{E_FAIL, VARIANT_BOOL};
use windows::Win32::System::Com::*;
use windows::Win32::System::Variant::*;

// ============================================================
// ComRuntime — RAII wrapper for COM STA initialization
// ============================================================

pub struct ComRuntime;

impl ComRuntime {
    pub fn init() -> Result<Self> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()? };
        Ok(Self)
    }
}

impl Drop for ComRuntime {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

// ============================================================
// Variant — Rust-friendly COM variant value
// ============================================================

#[derive(Debug, Clone)]
pub enum Variant {
    Empty,
    String(String),
    I32(i32),
    F64(f64),
    Bool(bool),
    Dispatch(DispatchObject),
}

impl Variant {
    pub fn into_dispatch(self) -> Result<DispatchObject> {
        match self {
            Variant::Dispatch(d) => Ok(d),
            _ => Err(Error::new(E_FAIL, "Variant is not a Dispatch object")),
        }
    }

    pub fn as_str(&self) -> Option<&str> {
        match self {
            Variant::String(s) => Some(s),
            _ => None,
        }
    }

    pub fn as_f64(&self) -> Option<f64> {
        match self {
            Variant::F64(n) => Some(*n),
            Variant::I32(n) => Some(*n as f64),
            _ => None,
        }
    }

    pub fn as_i32(&self) -> Option<i32> {
        match self {
            Variant::I32(n) => Some(*n),
            _ => None,
        }
    }

    pub fn as_bool(&self) -> Option<bool> {
        match self {
            Variant::Bool(b) => Some(*b),
            _ => None,
        }
    }

    pub fn is_empty(&self) -> bool {
        matches!(self, Variant::Empty)
    }
}

impl fmt::Display for Variant {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Variant::Empty => write!(f, "(empty)"),
            Variant::String(s) => write!(f, "{s}"),
            Variant::I32(n) => write!(f, "{n}"),
            Variant::F64(n) => write!(f, "{n}"),
            Variant::Bool(b) => write!(f, "{b}"),
            Variant::Dispatch(_) => write!(f, "(dispatch)"),
        }
    }
}

impl From<&str> for Variant {
    fn from(s: &str) -> Self {
        Variant::String(s.to_string())
    }
}
impl From<String> for Variant {
    fn from(s: String) -> Self {
        Variant::String(s)
    }
}
impl From<i32> for Variant {
    fn from(n: i32) -> Self {
        Variant::I32(n)
    }
}
impl From<f64> for Variant {
    fn from(n: f64) -> Self {
        Variant::F64(n)
    }
}
impl From<bool> for Variant {
    fn from(b: bool) -> Self {
        Variant::Bool(b)
    }
}

// ============================================================
// DispatchObject — IDispatch wrapper for late-bound COM
// ============================================================

#[derive(Clone)]
pub struct DispatchObject {
    inner: IDispatch,
}

impl fmt::Debug for DispatchObject {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("DispatchObject").finish()
    }
}

impl DispatchObject {
    /// Wrap an existing IDispatch
    pub fn from_raw(disp: IDispatch) -> Self {
        Self { inner: disp }
    }

    /// Create a new COM object from a ProgID (e.g. "Excel.Application")
    pub fn create(prog_id: &str) -> Result<Self> {
        unsafe {
            let wide: Vec<u16> = prog_id.encode_utf16().chain(std::iter::once(0)).collect();
            let clsid = CLSIDFromProgID(PCWSTR(wide.as_ptr()))?;
            let disp: IDispatch = CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)?;
            Ok(Self { inner: disp })
        }
    }

    /// Get a property (no args)
    pub fn get(&self, name: &str) -> Result<Variant> {
        self.invoke_raw(name, DISPATCH_PROPERTYGET, &[])
    }

    /// Get a parameterized property (e.g. Cells(row, col))
    pub fn get_by(&self, name: &str, args: &[Variant]) -> Result<Variant> {
        self.invoke_raw(name, DISPATCH_PROPERTYGET, args)
    }

    /// Set a property
    pub fn put(&self, name: &str, value: impl Into<Variant>) -> Result<()> {
        let v = value.into();
        self.invoke_put(name, &v)
    }

    /// Call a method
    pub fn call(&self, name: &str, args: &[Variant]) -> Result<Variant> {
        self.invoke_raw(
            name,
            DISPATCH_FLAGS(DISPATCH_METHOD.0 | DISPATCH_PROPERTYGET.0),
            args,
        )
    }

    // -- internals --

    fn get_dispid(&self, name: &str) -> Result<i32> {
        unsafe {
            let wide: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
            let name_ptr = PCWSTR(wide.as_ptr());
            let names = [name_ptr];
            let mut dispid = 0i32;
            self.inner
                .GetIDsOfNames(&GUID::zeroed(), names.as_ptr(), 1, 0, &mut dispid)?;
            Ok(dispid)
        }
    }

    fn invoke_raw(
        &self,
        name: &str,
        flags: DISPATCH_FLAGS,
        args: &[Variant],
    ) -> Result<Variant> {
        unsafe {
            let dispid = self.get_dispid(name)?;

            // COM expects args in reverse order
            let mut raw_args: Vec<VARIANT> = args.iter().rev().map(variant_to_raw).collect();

            let params = DISPPARAMS {
                rgvarg: if raw_args.is_empty() {
                    std::ptr::null_mut()
                } else {
                    raw_args.as_mut_ptr()
                },
                rgdispidNamedArgs: std::ptr::null_mut(),
                cArgs: raw_args.len() as u32,
                cNamedArgs: 0,
            };

            let mut result = VARIANT::default();
            let mut excep = EXCEPINFO::default();

            self.inner.Invoke(
                dispid,
                &GUID::zeroed(),
                0,
                flags,
                &params,
                Some(&mut result),
                Some(&mut excep),
                None,
            )?;

            Ok(variant_from_raw(&result))
        }
    }

    fn invoke_put(&self, name: &str, value: &Variant) -> Result<()> {
        unsafe {
            let dispid = self.get_dispid(name)?;

            let mut raw_value = variant_to_raw(value);
            let mut named_arg: i32 = -3; // DISPID_PROPERTYPUT

            let params = DISPPARAMS {
                rgvarg: &mut raw_value,
                rgdispidNamedArgs: &mut named_arg,
                cArgs: 1,
                cNamedArgs: 1,
            };

            let mut excep = EXCEPINFO::default();

            self.inner.Invoke(
                dispid,
                &GUID::zeroed(),
                0,
                DISPATCH_PROPERTYPUT,
                &params,
                None,
                Some(&mut excep),
                None,
            )?;

            Ok(())
        }
    }
}

// ============================================================
// VARIANT conversion helpers (raw union access)
// ============================================================

fn make_variant(vt: VARENUM, data: VARIANT_0_0_0) -> VARIANT {
    VARIANT {
        Anonymous: VARIANT_0 {
            Anonymous: ManuallyDrop::new(VARIANT_0_0 {
                vt,
                wReserved1: 0,
                wReserved2: 0,
                wReserved3: 0,
                Anonymous: data,
            }),
        },
    }
}

fn variant_to_raw(v: &Variant) -> VARIANT {
    match v {
        Variant::Empty => VARIANT::default(),
        Variant::String(s) => make_variant(
            VT_BSTR,
            VARIANT_0_0_0 {
                bstrVal: ManuallyDrop::new(BSTR::from(s.as_str())),
            },
        ),
        Variant::I32(n) => make_variant(VT_I4, VARIANT_0_0_0 { lVal: *n }),
        Variant::F64(n) => make_variant(VT_R8, VARIANT_0_0_0 { dblVal: *n }),
        Variant::Bool(b) => make_variant(
            VT_BOOL,
            VARIANT_0_0_0 {
                boolVal: VARIANT_BOOL(if *b { -1 } else { 0 }),
            },
        ),
        Variant::Dispatch(d) => make_variant(
            VT_DISPATCH,
            VARIANT_0_0_0 {
                pdispVal: ManuallyDrop::new(Some(d.inner.clone())),
            },
        ),
    }
}

fn variant_from_raw(v: &VARIANT) -> Variant {
    unsafe {
        let vt = v.Anonymous.Anonymous.vt;
        match vt {
            VT_EMPTY | VT_NULL => Variant::Empty,
            VT_BSTR => {
                let bstr = &*v.Anonymous.Anonymous.Anonymous.bstrVal;
                Variant::String(bstr.to_string())
            }
            VT_I2 => Variant::I32(v.Anonymous.Anonymous.Anonymous.iVal as i32),
            VT_I4 | VT_INT => Variant::I32(v.Anonymous.Anonymous.Anonymous.lVal),
            VT_R4 => Variant::F64(v.Anonymous.Anonymous.Anonymous.fltVal as f64),
            VT_R8 => Variant::F64(v.Anonymous.Anonymous.Anonymous.dblVal),
            VT_BOOL => {
                let b = v.Anonymous.Anonymous.Anonymous.boolVal;
                Variant::Bool(b.0 != 0)
            }
            VT_DISPATCH => {
                let disp = &*v.Anonymous.Anonymous.Anonymous.pdispVal;
                match disp {
                    Some(d) => Variant::Dispatch(DispatchObject { inner: d.clone() }),
                    None => Variant::Empty,
                }
            }
            _ => Variant::Empty,
        }
    }
}
