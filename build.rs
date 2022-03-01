fn main()
{
    cxx_build::bridge("src/lib.rs")
        .cpp(true)
        .file("src/vbs.cpp")
        .compile("inline-vbs-demo");
    println!("cargo:rerun-if-changed=src/lib.rs");
    println!("cargo:rerun-if-changed=src/vbs.cpp");
    println!("cargo:rerun-if-changed=include/vbs.h");
}