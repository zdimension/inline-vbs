# inline-vbs

`inline-vbs` is a crate that allows you to embed VBScript, JScript and many other languages inside Rust code files. It uses
the [Active Scripting](https://docs.microsoft.com/en-us/archive/msdn-magazine/2000/december/active-scripting-apis-add-powerful-custom-debugging-to-your-script-hosting-app) COM APIs to dynamically parse and execute (optionally, evaluate) code.

![image](https://user-images.githubusercontent.com/4533568/212424549-7440814e-64b4-4deb-853f-b28531904670.png)


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
  See more examples in [tests/tests.rs](tests/tests.rs).

## Installation
Add this to your `Cargo.toml`:
```toml
[dependencies]
inline-vbs = "0.4.0"
```

**Important:** You need to have the MSVC Build Tools installed on your computer, as required by [cc](https://github.com/rust-lang/cc-rs).

### Language support

VBScript (`vbs!`) and JScript (`js!`) are available out of the box on 32-bit and 64-bit.

Other Active Scripting engines exist:
- Ruby (`ruby!`): ActiveScriptRuby [1.8 (tested, 32-bit only)](https://www.artonx.org/data/asr/ActiveRuby.msi)
  - [2.4 (32 or 64-bit) (untested!)](https://www.artonx.org/data/asr/), you need to change the CLSID in [src/vbs.cpp](src/vbs.cpp)
- Perl (`perl!`): [ActivePerl 5.20 (32-bit)](https://raw.githubusercontent.com/PengjieRen/LibSum/master/ActivePerl-5.20.2.2002-MSWin32-x86-64int-299195.msi)

Note: install an engine matching the bitness of your program; by default Rust on Windows builds 
64-bit programs, which can only use 64-bit libraries. If you want to use a 32-bit library, you
need to build your program with `--target i686-pc-windows-msvc`.

## Limitations
Many. 

## Motivation
N/A

## License
This project is licensed under either of
* Apache License, Version 2.0, ([LICENSE-APACHE](LICENSE-APACHE) or
  https://www.apache.org/licenses/LICENSE-2.0)
* MIT license ([LICENSE-MIT](LICENSE-MIT) or
  https://opensource.org/licenses/MIT)
  at your option.
