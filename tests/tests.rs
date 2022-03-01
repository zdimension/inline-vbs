#[cfg(test)]
mod tests {
    use inline_vbs::vbs;
    use inline_vbs_macros::vbs_raw;

    #[test]
    fn bidule()
    {
        assert!(vbs![MsgBox Now].is_ok());

        assert!(vbs_raw!(r#"
         Function Square(x As Integer)
                Return x * x
            End Function
        "#).is_ok());

        assert!(vbs![Dim firstname].is_ok());
        assert!(vbs![firstname="Sasha"].is_ok());

        assert!(vbs![MsgBox "Bonjour " & firstname].is_ok());

        assert!(vbs![MsgBox 1 / 0].is_err());

        assert!(vbs![MsgBox "bye"].is_ok());
    }
}
