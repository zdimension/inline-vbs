//! `inline-vbs` is a crate that allows you to embed VBScript code inside Rust code files. It uses
//! the [Active Scripting](https://docs.microsoft.com/en-us/archive/msdn-magazine/2000/december/active-scripting-apis-add-powerful-custom-debugging-to-your-script-hosting-app) COM APIs to dynamically parse and execute (optionally, evaluate) code.
//!
//! ### Basic usage
//! ```rust
//! use inline_vbs::*;
//!
//! fn main() {
//!     vbs![On Error Resume Next]; // tired of handling errors?
//!     vbs![MsgBox "Hello, world!"];
//!     if let Ok(Variant::String(str)) = vbs_!["VBScript" & " Rocks!"] {
//!         println!("{}", str);
//!     }
//! }
//! ```
//!
//! Macros:
//! * `vbs!` - Executes a statement
//! * `vbs_!` - Evaluates an expression
//! * `vbs_raw!` - Executes a statement (string input instead of tokens, use for multiline code)
//!
//! See more examples in [tests/tests.rs](tests/tests.rs)
//!
//! ### Limitations
//! Many
//!
//! ### Motivation
//! N/A
//!
//! ### License
//!
//!
//! This project is licensed under either of
//!
//!  * Apache License, Version 2.0, ([LICENSE-APACHE](LICENSE-APACHE) or
//!    https://www.apache.org/licenses/LICENSE-2.0)
//!  * MIT license ([LICENSE-MIT](LICENSE-MIT) or
//!    https://opensource.org/licenses/MIT)
//!
//! at your option.

mod variant;

use std::os::raw::c_char;

use winapi::um::oaidl::VARIANT;

pub use inline_vbs_macros::{vbs, vbs_, vbs_raw};
pub use crate::variant::Variant;
use crate::variant::{decode_variant};

type HRESULT = i32;

#[cxx::bridge]
mod ffi
{
    unsafe extern "C++"
    {
        include!("inline-vbs/include/vbs.h");

        pub fn init() -> i32;
        pub unsafe fn parse_wrapper(str: &str, output: *mut c_char) -> i32;
        pub fn error_to_string(error: i32) -> String;
    }
}

fn decode_hr<T>(hr: i32, val: T) -> Result<T, String>
{
    if hr != 0 { Err(ffi::error_to_string(hr)) } else { Ok(val) }
}

fn or_die(hr: HRESULT)
{
    if let Err(msg) = decode_hr(hr, ())
    {
        panic!("Internal VBS error: {}", msg);
    }
}

pub fn run_code(code: &str) -> Result<(), String>
{
    unsafe
        {
            or_die(ffi::init());
            let result = ffi::parse_wrapper(code, std::ptr::null_mut());
            decode_hr(result, ())
        }
}

pub fn run_expr(code: &str) -> Result<Variant, String>
{
    unsafe
        {
            or_die(ffi::init());
            let mut variant: VARIANT = std::mem::zeroed();
            let result = ffi::parse_wrapper(code, (&mut variant) as *mut VARIANT as *mut c_char);
            decode_hr(result, decode_variant(variant))
        }
}

