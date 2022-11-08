# inline-vbs

`inline-vbs` is a crate that allows you to embed VBScript code inside Rust code files. It uses
the [Active Scripting](https://docs.microsoft.com/en-us/archive/msdn-magazine/2000/december/active-scripting-apis-add-powerful-custom-debugging-to-your-script-hosting-app) COM APIs to dynamically parse and execute (optionally, evaluate) code.

## Basic usage
```rust
use inline_vbs::*;

fn main() {
    vbs! { On Error Resume Next } // tired of handling errors?
    vbs! { MsgBox "Hello, world!" }
    let language = "VBScript";
    assert_eq!(vbs_!['language & " Rocks!"], "VBScript Rocks!".into());
}
```
Macros:
* `vbs!` - Executes a statement or evaluates an expression (depending on context)
* `vbs_!` - Evaluates an expression
* `vbs_raw!` - Executes a statement (string input instead of tokens, use for multiline code)
  See more examples in [tests/tests.rs](tests/tests.rs)

## Installation
Add this to your `Cargo.toml`:
```toml
[dependencies]
inline-vbs = "0.2.1"
```

**Important:** You need to have the MSVC Build Tools installed on your computer, as required by [cc](https://github.com/rust-lang/cc-rs).

## Limitations
Many. Most notably, `IDispatch` objects (i.e. what `CreateObject` returns) can't be passed to
the engine (`let x = vbs! { CreateObject("WScript.Shell") }; vbs! { y = 'x }` won't work).

## Motivation
N/A

## License
This project is licensed under either of
* Apache License, Version 2.0, ([LICENSE-APACHE](LICENSE-APACHE) or
  https://www.apache.org/licenses/LICENSE-2.0)
* MIT license ([LICENSE-MIT](LICENSE-MIT) or
  https://opensource.org/licenses/MIT)
  at your option.