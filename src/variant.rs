use winapi::shared::wtypes::VARTYPE;
use winapi::um::oaidl::VARIANT;
use enumn::N;
use winapi::um::winnt::HRESULT;

#[derive(PartialEq, Debug)]
pub enum Variant
{
    None,
    Bool(bool),
    I8(i8),
    I16(i16),
    I32(i32),
    I64(i64),
    U8(u8),
    U16(u16),
    U32(u32),
    U64(u64),
    F32(f32),
    F64(f64),
    String(String),
    Error(HRESULT),
    Unimplemented(VariantType),
    Unknown(VARTYPE),
}

#[derive(N, Debug, PartialEq)]
#[allow(non_camel_case_types)]
pub enum VariantType
{
    VT_EMPTY = 0,
    VT_NULL = 1,
    VT_I2 = 2,
    VT_I4 = 3,
    VT_R4 = 4,
    VT_R8 = 5,
    VT_CY = 6,
    VT_DATE = 7,
    VT_BSTR = 8,
    VT_DISPATCH = 9,
    VT_ERROR = 10,
    VT_BOOL = 11,
    VT_VARIANT = 12,
    VT_UNKNOWN = 13,
    VT_DECIMAL = 14,
    VT_I1 = 16,
    VT_UI1 = 17,
    VT_UI2 = 18,
    VT_UI4 = 19,
    VT_I8 = 20,
    VT_UI8 = 21,
    VT_INT = 22,
    VT_UINT = 23,
    VT_VOID = 24,
    VT_HRESULT = 25,
    VT_PTR = 26,
    VT_SAFEARRAY = 27,
    VT_CARRAY = 28,
    VT_USERDEFINED = 29,
    VT_LPSTR = 30,
    VT_LPWSTR = 31,
    VT_RECORD = 36,
    VT_INT_PTR = 37,
    VT_UINT_PTR = 38,
}

pub fn decode_variant(v: VARIANT) -> Variant
{
    unsafe {
        let val = v.n1.n2();
        match VariantType::n(val.vt)
        {
            Some(VariantType::VT_EMPTY) => Variant::None,
            Some(VariantType::VT_BOOL) => Variant::Bool(*val.n3.boolVal() != 0),
            Some(VariantType::VT_I1) => Variant::I8(*val.n3.cVal()),
            Some(VariantType::VT_I2) => Variant::I16(*val.n3.iVal()),
            Some(VariantType::VT_I4) => Variant::I32(*val.n3.lVal()),
            Some(VariantType::VT_I8) => Variant::I64(*val.n3.llVal()),
            Some(VariantType::VT_UI1) => Variant::U8(*val.n3.bVal()),
            Some(VariantType::VT_UI2) => Variant::U16(*val.n3.uiVal()),
            Some(VariantType::VT_UI4) => Variant::U32(*val.n3.ulVal()),
            Some(VariantType::VT_UI8) => Variant::U64(*val.n3.ullVal()),
            Some(VariantType::VT_R4) => Variant::F32(*val.n3.fltVal()),
            Some(VariantType::VT_R8) => Variant::F64(*val.n3.dblVal()),
            Some(VariantType::VT_BSTR) => Variant::String(widestring::U16CString::from_ptr_str(*val.n3.bstrVal()).to_string().unwrap()),
            Some(VariantType::VT_ERROR) => Variant::Error(*val.n3.scode()),
            Some(t) => Variant::Unimplemented(t),
            None => Variant::Unknown(val.vt),
        }
    }
}
