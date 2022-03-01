# embed-c

`inline-vbs` is a crate that allows you to embed VBScript code inside Rust code files. It uses
the [Active Scripting](https://docs.microsoft.com/en-us/archive/msdn-magazine/2000/december/active-scripting-apis-add-powerful-custom-debugging-to-your-script-hosting-app) COM APIs to dynamically parse and execute (optionally, evaluate) code.

## Basic usage
```rust
use inline_vbs::*;

fn main() {
    vbs![On Error Resume Next]; // tired of handling errors?
    vbs![MsgBox "Hello, world!"];
    if let Ok(Variant::String(str)) = vbs_!["VBScript" & " Rocks!"] {
        println!("{}", str);
    }
}
```

## Installation
Add this to your `Cargo.toml`:
```toml
[dependencies]
inline-vbs = "0.1"
```

**Important:** You need to have the MSVC Build Tools installed on your computer, and you may need to run
the `vsdevcmd.bat` script in your terminal to set up the build environment:
```
"C:\Program Files\Microsoft Visual Studio\2022\Enterprise\Common7\Tools\VsDevCmd.bat" -arch=amd64
```
(replace with the path of your VS installation)

## Limitations
Many

## Motivation
N/A

## License
This project is licensed under either of
* Apache License, Version 2.0, ([LICENSE-APACHE](LICENSE-APACHE) or
  https://www.apache.org/licenses/LICENSE-2.0)
* MIT license ([LICENSE-MIT](LICENSE-MIT) or
  https://opensource.org/licenses/MIT)
  at your option.