// Caps linear memory at 1 GiB, so a hostile workbook fails inside the engine instead of
// growing the host process; Po's 25 MiB file limit keeps real workbooks far below it.
fn main() {
    println!("cargo:rustc-link-arg-cdylib=--max-memory=1073741824");
}
