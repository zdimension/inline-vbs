type HRESULT = i32;

#[cxx::bridge]
mod ffi
{
    unsafe extern "C++"
    {
        include!("inline-vbs/include/vbs.h");

        pub fn init() -> i32;
        pub fn parse(str: &str) -> i32;
        pub fn close() -> i32;
        pub fn error_to_string(error: i32) -> String;
    }
}

unsafe fn decode_hr(hr: i32) -> Result<(), String>
{
    if hr != 0 { Err(ffi::error_to_string(hr)) } else { Ok(()) }
}

unsafe fn or_die(hr: HRESULT)
{
    if let Err(msg) = decode_hr(hr)
    {
        panic!("Internal VBS error: {}", msg);
    }
}

pub fn run_code(code: &str) -> Result<(), String>
{
    unsafe
        {
            or_die(ffi::init());
            let result = ffi::parse(code);
            // todo: maybe handle multiple contexts, one day?
            decode_hr(result)
        }
}

pub use inline_vbs_macros::vbs;
