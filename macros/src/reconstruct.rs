/*
Copyright 2020- Ivan Enderlin

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

1. Redistributions of source code must retain the above copyright
   notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright
   notice, this list of conditions and the following disclaimer in the
   documentation and/or other materials provided with the
   distribution.

3. Neither the name of the copyright holder nor the names of its
   contributors may be used to endorse or promote products derived
   from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

use proc_macro2::TokenStream;

pub fn reconstruct(input: TokenStream) -> String {
    use proc_macro2::{Delimiter, Spacing, TokenTree::*};

    let mut output = String::new();
    let mut iterator = input.into_iter().peekable();

    loop {
        match iterator.next() {
            Some(Punct(token)) => {
                let token_value = token.as_char();

                match token_value {
                    _ => {
                        output.push(token_value);

                        if token.spacing() == Spacing::Alone {
                            output.push(' ');
                        }
                    }
                }
            }

            Some(Ident(ident)) => {
                output.push_str(&ident.to_string());
                output.push(' ');
            }

            Some(Group(group)) => {
                let group_output = reconstruct(group.stream());

                match group.delimiter() {
                    Delimiter::Parenthesis => {
                        output.push('(');
                        output.push_str(&group_output);
                        output.push(')');
                    }

                    Delimiter::Brace => {
                        output.push('{');
                        output.push('\n');
                        output.push_str(&group_output);
                        output.push('\n');
                        output.push('}');
                    }

                    Delimiter::Bracket => {
                        output.push('[');
                        output.push_str(&group_output);
                        output.push(']');
                    }

                    Delimiter::None => {
                        output.push_str(&group_output);
                    }
                }
            }

            Some(token) => {
                output.push_str(&token.to_string());
            }

            None => break,
        }
    }

    output
}