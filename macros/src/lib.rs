mod reconstruct;

use litrs::StringLit;
use proc_macro2::{TokenStream, TokenTree};
use quote::quote;
extern crate proc_macro;

#[proc_macro]
pub fn vbs(input: proc_macro::TokenStream) -> proc_macro::TokenStream
{
    use crate::reconstruct::reconstruct;
    let code = reconstruct(TokenStream::from(input));

    println!("Running: {}", code);

    quote! [::inline_vbs::run_code(#code)].into()
}

#[proc_macro]
pub fn vbs_raw(input: proc_macro::TokenStream) -> proc_macro::TokenStream
{
    let input = input.into_iter().collect::<Vec<_>>();

    if input.len() != 1 {
        let msg = format!("expected exactly one input token, got {}", input.len());
        return quote! { compile_error!(#msg) }.into();
    }

    let string_lit = match StringLit::try_from(&input[0]) {
        Err(e) => return e.to_compile_error(),
        Ok(lit) => lit,
    };

    let code = string_lit.value();

    quote! [::inline_vbs::run_code(#code)].into()
}