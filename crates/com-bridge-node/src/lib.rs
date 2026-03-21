#![deny(clippy::all)]

use std::sync::Once;

use com_core::{ComRuntime, DispatchObject, MemberKind, Variant};
use napi::bindgen_prelude::*;
use napi::{JsUnknown, NapiRaw};
use napi_derive::napi;

// COM runtime: 프로세스 당 1회 초기화, 영구 유지
static COM_INIT: Once = Once::new();

fn ensure_init() {
    COM_INIT.call_once(|| {
        let rt = ComRuntime::init().expect("COM STA init failed");
        std::mem::forget(rt); // COM은 프로세스 종료까지 유지
    });
}

fn to_napi_err(e: windows::core::Error) -> Error {
    Error::from_reason(format!("COM error: {e}"))
}

// ============================================================
// Lifecycle
// ============================================================

#[napi]
pub fn com_init() {
    ensure_init();
}

#[napi]
pub fn com_create(prog_id: String) -> Result<External<DispatchObject>> {
    ensure_init();
    let disp = DispatchObject::create(&prog_id).map_err(to_napi_err)?;
    Ok(External::new(disp))
}

// ============================================================
// Generic dispatch — use #[js_function] for dynamic types
// ============================================================

#[napi(ts_return_type = "any")]
pub fn com_get(env: Env, handle: External<DispatchObject>, prop: String, args: Option<Vec<String>>) -> Result<JsUnknown> {
    let disp = &*handle;
    let result = match args {
        Some(ref str_args) if !str_args.is_empty() => {
            let variants: Vec<Variant> = str_args.iter().map(|s| parse_variant(s)).collect();
            disp.get_by(&prop, &variants).map_err(to_napi_err)?
        }
        _ => disp.get(&prop).map_err(to_napi_err)?,
    };
    variant_to_js(&env, result)
}

#[napi(ts_return_type = "any")]
pub fn com_get_by(env: Env, handle: External<DispatchObject>, prop: String, args: Vec<External<DispatchObject>>) -> Result<JsUnknown> {
    // Overload for dispatch-type args (e.g., passing COM objects as parameters)
    let disp = &*handle;
    let variants: Vec<Variant> = args.iter().map(|a| Variant::Dispatch((**a).clone())).collect();
    let result = disp.get_by(&prop, &variants).map_err(to_napi_err)?;
    variant_to_js(&env, result)
}

#[napi]
pub fn com_put_str(handle: External<DispatchObject>, prop: String, value: String) -> Result<()> {
    let disp = &*handle;
    disp.put(&prop, Variant::String(value)).map_err(to_napi_err)
}

#[napi]
pub fn com_put_num(handle: External<DispatchObject>, prop: String, value: f64) -> Result<()> {
    let disp = &*handle;
    let v = if value.fract() == 0.0 && value >= i32::MIN as f64 && value <= i32::MAX as f64 {
        Variant::I32(value as i32)
    } else {
        Variant::F64(value)
    };
    disp.put(&prop, v).map_err(to_napi_err)
}

#[napi]
pub fn com_put_bool(handle: External<DispatchObject>, prop: String, value: bool) -> Result<()> {
    let disp = &*handle;
    disp.put(&prop, Variant::Bool(value)).map_err(to_napi_err)
}

#[napi]
pub fn com_put_dispatch(handle: External<DispatchObject>, prop: String, value: External<DispatchObject>) -> Result<()> {
    let disp = &*handle;
    disp.put(&prop, Variant::Dispatch((*value).clone())).map_err(to_napi_err)
}

/// Universal put — auto-detects JS type
#[napi(ts_args_type = "handle: ExternalObject<unknown>, prop: string, value: any")]
pub fn com_put(env: Env, handle: External<DispatchObject>, prop: String, value: JsUnknown) -> Result<()> {
    let disp = &*handle;
    let variant = js_unknown_to_variant(&env, value)?;
    disp.put(&prop, variant).map_err(to_napi_err)
}

#[napi(ts_return_type = "any")]
pub fn com_call(env: Env, handle: External<DispatchObject>, method: String, args: Option<Vec<String>>) -> Result<JsUnknown> {
    let disp = &*handle;
    let variants: Vec<Variant> = match args {
        Some(ref str_args) => str_args.iter().map(|s| parse_variant(s)).collect(),
        None => vec![],
    };
    let result = disp.call(&method, &variants).map_err(to_napi_err)?;
    variant_to_js(&env, result)
}

/// Call with mixed typed args (handles JsUnknown[])
#[napi(ts_args_type = "handle: ExternalObject<unknown>, method: string, args: any[]", ts_return_type = "any")]
pub fn com_call_with(env: Env, handle: External<DispatchObject>, method: String, args: Vec<JsUnknown>) -> Result<JsUnknown> {
    let disp = &*handle;
    let variants: Vec<Variant> = args
        .into_iter()
        .map(|v| js_unknown_to_variant(&env, v))
        .collect::<Result<_>>()?;
    let result = disp.call(&method, &variants).map_err(to_napi_err)?;
    variant_to_js(&env, result)
}

// ============================================================
// Introspection
// ============================================================

#[napi(object)]
pub struct JsMemberInfo {
    pub name: String,
    pub kind: String,
    pub params: Vec<JsParamInfo>,
    pub return_type: String,
}

#[napi(object)]
pub struct JsParamInfo {
    pub name: String,
    pub vt: String,
}

#[napi]
pub fn com_list_members(handle: External<DispatchObject>) -> Result<Vec<JsMemberInfo>> {
    let disp = &*handle;
    let members = disp.list_members().map_err(to_napi_err)?;
    Ok(members
        .into_iter()
        .map(|m| JsMemberInfo {
            name: m.name,
            kind: match m.kind {
                MemberKind::Method => "method".into(),
                MemberKind::Get => "get".into(),
                MemberKind::Put => "put".into(),
                MemberKind::GetPut => "get/put".into(),
            },
            params: m.params.into_iter().map(|p| JsParamInfo { name: p.name, vt: p.vt }).collect(),
            return_type: m.return_type,
        })
        .collect())
}

// ============================================================
// JS ↔ Variant conversion
// ============================================================

/// Parse string to Variant (simple heuristic, for com_get/com_call string args)
fn parse_variant(s: &str) -> Variant {
    if let Ok(n) = s.parse::<i32>() {
        return Variant::I32(n);
    }
    if let Ok(n) = s.parse::<f64>() {
        return Variant::F64(n);
    }
    match s {
        "true" => Variant::Bool(true),
        "false" => Variant::Bool(false),
        "" => Variant::Empty,
        _ => Variant::String(s.to_string()),
    }
}

/// Convert JsUnknown to Variant (for com_put / com_call_with)
fn js_unknown_to_variant(env: &Env, val: JsUnknown) -> Result<Variant> {
    let vtype = val.get_type()?;
    match vtype {
        ValueType::Undefined | ValueType::Null => Ok(Variant::Empty),
        ValueType::Boolean => {
            let b = val.coerce_to_bool()?.get_value()?;
            Ok(Variant::Bool(b))
        }
        ValueType::Number => {
            let n = val.coerce_to_number()?.get_double()?;
            if n.fract() == 0.0 && n >= i32::MIN as f64 && n <= i32::MAX as f64 {
                Ok(Variant::I32(n as i32))
            } else {
                Ok(Variant::F64(n))
            }
        }
        ValueType::String => {
            let s = val.coerce_to_string()?.into_utf8()?.into_owned()?;
            Ok(Variant::String(s))
        }
        ValueType::External => {
            // Extract DispatchObject from External handle
            let external = unsafe {
                External::<DispatchObject>::from_napi_value(env.raw(), val.raw())?
            };
            Ok(Variant::Dispatch((*external).clone()))
        }
        _ => Err(Error::from_reason("Unsupported JS type for COM variant")),
    }
}

/// Convert Variant to JS value
fn variant_to_js(env: &Env, val: Variant) -> Result<JsUnknown> {
    match val {
        Variant::Empty => env.get_null().map(|v| v.into_unknown()),
        Variant::String(s) => env.create_string(&s).map(|v| v.into_unknown()),
        Variant::I32(n) => env.create_int32(n).map(|v| v.into_unknown()),
        Variant::F64(n) => env.create_double(n).map(|v| v.into_unknown()),
        Variant::Bool(b) => env.get_boolean(b).map(|v| v.into_unknown()),
        Variant::Dispatch(disp) => {
            let external = env.create_external(disp, None)?;
            Ok(external.into_unknown())
        }
    }
}
