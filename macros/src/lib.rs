/*
Includes bits of code from the great & almighty Mara Bos and her amazing inline-python project.

Copyright (c) 2019-2020 Drones for Work B.V. and contributors
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.
2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR
ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

#![feature(proc_macro_span)]

use litrs::StringLit;
use proc_macro::LineColumn;
use proc_macro::Span;
use proc_macro2::TokenStream;
use proc_macro2::{Delimiter, Ident, Spacing, TokenTree};
use quote::quote;
use quote::quote_spanned;
use std::collections::BTreeMap;
use std::fmt::Write;

extern crate proc_macro;

fn macro_impl(input: TokenStream) -> Result<TokenStream, TokenStream> {
    let mut x = EmbedVbs::new();

    x.add(input)?;

    let EmbedVbs {
        code, variables, ..
    } = x;

    let varname = variables.keys();
    let var = variables.values();

    Ok(quote! {
        {
            #(::inline_vbs::set_variable(#varname, #var).unwrap();)*
            ::inline_vbs::Runner::run_code(#code)
        }
    })
}

#[proc_macro]
pub fn vbs(input: proc_macro::TokenStream) -> proc_macro::TokenStream {
    proc_macro::TokenStream::from(process_tokens(input))
}

fn process_tokens(input: proc_macro::TokenStream) -> TokenStream {
    match macro_impl(TokenStream::from(input)) {
        Ok(tokens) => tokens,
        Err(tokens) => tokens,
    }
}

#[proc_macro]
pub fn vbs_(input: proc_macro::TokenStream) -> proc_macro::TokenStream {
    let call = process_tokens(input);
    (quote! {
        {
            let res: Variant = #call;
            res
        }
    })
    .into()
}

#[proc_macro]
pub fn vbs_raw(input: proc_macro::TokenStream) -> proc_macro::TokenStream {
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

struct EmbedVbs {
    pub code: String,
    pub variables: BTreeMap<String, Ident>,
    pub first_indent: Option<usize>,
    pub loc: LineColumn,
}

impl EmbedVbs {
    pub fn new() -> Self {
        Self {
            code: String::new(),
            variables: BTreeMap::new(),
            loc: LineColumn { line: 1, column: 0 },
            first_indent: None,
        }
    }

    fn add_whitespace(&mut self, span: Span, loc: LineColumn) -> Result<(), TokenStream> {
        #[allow(clippy::comparison_chain)]
        if loc.line > self.loc.line {
            while loc.line > self.loc.line {
                self.code.push('\n');
                self.loc.line += 1;
            }
            let first_indent = *self.first_indent.get_or_insert(loc.column);
            let indent = loc.column.checked_sub(first_indent);
            let indent = indent.ok_or_else(
                || quote_spanned!(span.into() => compile_error!{"Invalid indentation"}),
            )?;
            for _ in 0..indent {
                self.code.push(' ');
            }
            self.loc.column = loc.column;
        } else if loc.line == self.loc.line {
            while loc.column > self.loc.column {
                self.code.push(' ');
                self.loc.column += 1;
            }
        }

        Ok(())
    }

    pub fn add(&mut self, input: TokenStream) -> Result<(), TokenStream> {
        let mut tokens = input.into_iter();

        while let Some(token) = tokens.next() {
            let span = token.span().unwrap();
            self.add_whitespace(span, span.start())?;

            match &token {
                TokenTree::Group(x) => {
                    let (start, end) = match x.delimiter() {
                        Delimiter::Parenthesis => ("(", ")"),
                        Delimiter::Brace => ("{", "}"),
                        Delimiter::Bracket => ("[", "]"),
                        Delimiter::None => ("", ""),
                    };
                    self.code.push_str(start);
                    self.loc.column += start.len();
                    self.add(x.stream())?;
                    let mut end_loc = token.span().unwrap().end();
                    end_loc.column = end_loc.column.saturating_sub(end.len());
                    self.add_whitespace(span, end_loc)?;
                    self.code.push_str(end);
                    self.loc.column += end.len();
                }
                TokenTree::Punct(x) => {
                    if x.as_char() == '\'' && x.spacing() == Spacing::Joint {
                        let name = if let Some(TokenTree::Ident(name)) = tokens.next() {
                            name
                        } else {
                            unreachable!()
                        };
                        let name_str = format!("RUST{}", name);
                        self.code.push_str(&name_str);
                        self.loc.column += name_str.chars().count() - 4 + 1;
                        self.variables.entry(name_str).or_insert(name);
                    } else {
                        self.code.push(x.as_char());
                        self.loc.column += 1;
                    }
                }
                TokenTree::Ident(x) => {
                    write!(&mut self.code, "{}", x).unwrap();
                    self.loc = token.span().unwrap().end();
                }
                TokenTree::Literal(x) => {
                    let s = x.to_string();
                    self.code += &s;
                    self.loc = token.span().unwrap().end();
                }
            }
        }

        Ok(())
    }
}
