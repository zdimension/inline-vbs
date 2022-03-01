#[cfg(test)]
mod tests {
    use inline_vbs::*;

    #[test]
    fn test_statements()
    {
        assert!(vbs_raw!(r#"
            Function Square(x)
                Square = x * x
            End Function
        "#).is_ok());

        assert!(vbs![Dim variable].is_ok());
        assert!(vbs![variable="Sasha"].is_ok());

        assert!(vbs![variable = "Bonjour " & variable & ", " & Square(5)].is_ok());

        assert!(vbs![variable = 1 / 0].is_err());

        assert!(vbs![variable = "bye"].is_ok());

        assert_eq!(Ok(Variant::I16(4)), vbs_![2 + 2]);

        assert_eq!(Ok(Variant::String("Hello123bye".to_string())), vbs_!["Hello" & 123 & variable]);
    }
}
