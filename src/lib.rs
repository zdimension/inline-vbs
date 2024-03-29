#![doc = include_str!("../README.md")]

use std::os::raw::c_char;
pub use variant_rs::variant::ToVariant;
pub use variant_rs::variant::Variant;

use windows::Win32::System::Com::VARIANT;
use windows::Win32::System::Ole::VariantInit;

pub use inline_vbs_macros::*;

#[allow(clippy::upper_case_acronyms)]
type HRESULT = i32;

#[cxx::bridge]
mod ffi {
    enum ScriptLang {
        VBScript,
        JScript,
        Perl,
        Ruby,
        Last,
    }

    unsafe extern "C++" {
        include!("inline-vbs/include/vbs.h");

        pub fn init() -> i32;
        pub unsafe fn parse_wrapper(str: &str, output: *mut c_char, lang: ScriptLang) -> i32;
        pub fn error_to_string(error: i32) -> String;
        pub unsafe fn set_variable(name: &str, val: *mut c_char, lang: ScriptLang) -> i32;
    }
}

pub use ffi::ScriptLang;

fn decode_hr<T>(hr: i32, val: T) -> Result<T, String> {
    if hr != 0 {
        print!("Error: 0x{:08x}\n", hr);
        Err(ffi::error_to_string(hr))
    } else {
        Ok(val)
    }
}

fn or_die(hr: HRESULT) {
    if let Err(msg) = decode_hr(hr, ()) {
        panic!("Internal VBS error: {}", msg);
    }
}

pub trait Runner {
    fn run_code(code: &str, lang: ScriptLang) -> Self;
}

impl Runner for Variant {
    fn run_code(code: &str, lang: ScriptLang) -> Self {
        unsafe {
            or_die(ffi::init());
            let mut variant = VariantInit();
            let result = ffi::parse_wrapper(
                code.trim(),
                (&mut variant) as *mut VARIANT as *mut c_char,
                lang,
            );
            decode_hr(result, ()).expect("Non-success HRESULT");
            variant.try_into().expect("Can't decode VARIANT")
        }
    }
}

impl Runner for () {
    fn run_code(code: &str, lang: ffi::ScriptLang) -> Self {
        unsafe {
            or_die(ffi::init());
            let result = ffi::parse_wrapper(code, std::ptr::null_mut(), lang);
            decode_hr(result, ()).expect("Non-success HRESULT");
        }
    }
}

pub fn set_variable(name: &str, val: impl ToVariant, lang: ScriptLang) -> Result<(), String> {
    unsafe {
        //let _: () = Runner::run_code(format!("Dim {}", name).as_str());
        or_die(ffi::init());
        let var = val.to_variant();
        let ptr: VARIANT = var
            .try_into()
            .map_err(|e| format!("Variant conversion error: {:?}", e))?;
        let result = ffi::set_variable(name, &ptr as *const VARIANT as *mut _, lang);
        decode_hr(result, ())
    }
}
