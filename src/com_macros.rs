/// Generates a COM property getter for a method like `get_PropertyName`.
#[macro_export]
macro_rules! com_get {
    ($self:ident, $method:ident, $ty:ty) => {{
        unsafe {
            let mut val: $ty = std::mem::zeroed();
            let hr = ((*$self.ptr).lpVtbl.as_ref().unwrap().$method.unwrap())($self.ptr, &mut val);
            windows::core::HRESULT(hr).ok()?;
            Ok(val)
        }
    }};
}

/// Generates a COM property setter for a method like `put_PropertyName`.
#[macro_export]
macro_rules! com_put {
    ($self:ident, $method:ident, $val:expr) => {{
        unsafe {
            match $val {
                Some(v) => {
                    let hr = ((*$self.ptr).lpVtbl.as_ref().unwrap().$method.unwrap())($self.ptr, v);
                    windows::core::HRESULT(hr).ok()
                }
                None => Err(windows::core::Error::new(
                    windows::core::HRESULT(0x80004003u32 as i32), // E_POINTER
                    "Attempted to set None on non-nullable COM property".to_string(),
                )),
            }
        }
    }};
}

/// Generates a COM BSTR getter (returns String).
#[macro_export]
macro_rules! com_get_bstr {
    ($self:ident, $method:ident) => {{
        unsafe {
            let mut raw: *mut u16 = std::ptr::null_mut();
            let hr = ((*$self.ptr).lpVtbl.as_ref().unwrap().$method.unwrap())($self.ptr, &mut raw);
            windows::core::HRESULT(hr).ok()?;
            let bstr = windows::core::BSTR::from_raw(raw);
            Ok(bstr.to_string())
        }
    }};
}

/// Generates a COM BSTR setter (takes &str).
#[macro_export]
macro_rules! com_put_bstr {
    ($self:ident, $method:ident, $val:expr) => {{
        unsafe {
            let raw = match $val {
                Some(s) => windows::core::BSTR::from(s).as_ptr() as *mut _,
                None => std::ptr::null_mut(),
            };
            let hr = ((*$self.ptr).lpVtbl.as_ref().unwrap().$method.unwrap())($self.ptr, raw);
            windows::core::HRESULT(hr).ok()
        }
    }};
}
