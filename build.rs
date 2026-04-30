use std::env;
use std::fs;
use std::path::PathBuf;

fn main() {
    // Copy acronyms.csv (repo root) into the Cargo output directory so that
    // `cargo run` finds the file next to the compiled executable.
    let out_dir = PathBuf::from(env::var("OUT_DIR").unwrap());
    // OUT_DIR is <profile>/build/<crate>/out – the exe lives two levels up.
    let exe_dir = out_dir
        .ancestors()
        .nth(3)
        .expect("unexpected OUT_DIR depth")
        .to_path_buf();

    let src = PathBuf::from("acronyms.csv");
    let dst = exe_dir.join("acronyms.csv");

    if src.exists() && !dst.exists() {
        fs::copy(&src, &dst).expect("failed to copy acronyms.csv");
    }

    // Re-run only when the source file changes.
    println!("cargo:rerun-if-changed=acronyms.csv");
}
