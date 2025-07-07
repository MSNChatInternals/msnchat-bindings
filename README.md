# 🧩 msnchat-bindings

Rust FFI bindings for the legacy **MSN Chat ActiveX control**, reverse-engineered from the original `MSNChat45.ocx`. This crate provides low-level access to the COM interfaces and types used by MSN Chat 4.5-era chat components.

> ⚠️ These bindings are raw and unsafe by nature. They are intended for developers working with legacy COM automation, reverse engineering, or historical software preservation.

---

## ✨ Features

- ✅ Raw FFI bindings generated via [`bindgen`](https://github.com/rust-lang/rust-bindgen)
- ✅ Compatible with the original `MSNChat45.ocx` ActiveX control
- ✅ Includes COM interface definitions, enums, and constants
- 🧪 Ideal for experimentation, automation, or building a safe wrapper layer

---

## 📦 Usage

Add this crate to your `Cargo.toml`:

```
[dependencies]
msnchat-bindings = { git = "https://github.com/msnchatinternals/msnchat-bindings.git", branch = "main" }
```

Then in your code:

```
use msnchat_bindings::bindings::*;
```

You can now call into the raw COM interfaces or exported functions defined in the original OCX.

---

## 🛠 Requirements

- Rust 1.70+ recommended
- Windows SDK (for building and linking against COM types)
- The original `MSNChat45.ocx` must be registered via `regsvr32` if you intend to instantiate the control

---

## 🧠 Example (COM Instantiation)

```
use windows::Win32::System::Com::{
    CoInitializeEx, CoCreateInstance, CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED,
};
use windows::core::GUID;
use msnchat_bindings::bindings::IMSNAddin;

unsafe {
    CoInitializeEx(std::ptr::null_mut(), COINIT_APARTMENTTHREADED).unwrap();

    let clsid = GUID::from("YOUR-CLSID-HERE");
    let mut obj = std::ptr::null_mut();
    CoCreateInstance(
        &clsid,
        None,
        CLSCTX_INPROC_SERVER,
        &IMSNAddin::IID,
        &mut obj as *mut _ as _,
    )
    .unwrap();

    let addin = obj as *mut IMSNAddin;
    ((*(*addin).lpVtbl).OnLogon)(addin /*, args */);
}
```

---

## 🧬 Project Goals

- Preserve and document legacy COM interfaces from MSN Chat
- Enable safe and idiomatic Rust wrappers over time
- Support historical software exploration and compatibility tooling

---

## 📜 License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

---

## 🙏 Acknowledgments

- Microsoft, for the original MSN Chat platform
- The Rust FFI and reverse engineering communities
- [`bindgen`](https://github.com/rust-lang/rust-bindgen) and [`windows`](https://github.com/microsoft/windows-rs)

---

## 💬 Questions or Contributions?

Feel free to open an issue or pull request if you’d like to contribute, report bugs, or share ideas for safe wrappers and extensions.
