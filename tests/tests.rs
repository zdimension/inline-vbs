#[cfg(test)]
mod tests {
    use inline_vbs::*;
    use variant_rs::call;

    #[test]
    fn test_statements() {
        ruby! {
            def strrb
                return "Hello from Ruby"
            end
            def wsh
                wsh = WIN32OLE.new("WScript.Shell")
                return wsh
            end
            wsh.Popup(strrb)
        }
        perl! {
            sub strpl {
                return "@{[$Ruby->strrb()]} and Perl";
            }
            $Ruby->wsh->Popup(strpl());
        }
        let strrs = format!("{} and Rust", perl_! { strpl }.expect_string());
        let wsh = js_! { new ActiveXObject("WScript.Shell") }
            .expect_dispatch()
            .unwrap();
        call!(wsh, Popup(strrs)).unwrap();
        vbs! {
            strvb = 'strrs & " and VBScript"
            Ruby.wsh.Popup(strvb)
        }
        js! {
            var strjs = strvb + " and JScript";
            Ruby.wsh.Popup(strjs);
        }

        vbs! {
            Function Square(x)
                Square = x * x
            End Function
        }
        assert_eq!(vbs_!(Square(2)), Variant::I16(4));

        vbs! {
            variable = "Sasha"
        }

        assert_eq!(vbs_!(variable), "Sasha".into());

        let name = "inline_vbs";

        assert_eq!(vbs_!["Hello" & 123 & 'name], "Hello123inline_vbs".into());
    }
}
