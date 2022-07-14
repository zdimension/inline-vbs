#[cfg(test)]
mod tests {
    use inline_vbs::*;
    use variant_rs::variant::Variant;

    #[test]
    fn test_statements() {
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
