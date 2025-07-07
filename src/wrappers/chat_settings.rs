use windows::{
    Win32::System::{
        Com::{CLSCTX_INPROC_SERVER, CoCreateInstance},
        Ole::IOleObject,
    },
    core::{Interface, Result},
};

use crate::bindings::{
    guids::{CLSID_ChatSettings, IID_IChatSettings},
    ichat_settings::{IChatSettings, IChatSettingsVtbl},
};

pub struct ChatSettings {
    ptr: *mut IChatSettings,
}

impl ChatSettings {
    /// Constructs a `ChatSettings` from a raw COM pointer.
    pub fn from_raw(ptr: *mut IChatSettings) -> Self {
        Self { ptr }
    }

    /// Creates a new `ChatSettings` instance via `CoCreateInstance`.
    pub fn create() -> Result<Self> {
        // Step 1: Create the object using a known COM interface
        let ole: IOleObject =
            unsafe { CoCreateInstance(&CLSID_ChatSettings, None, CLSCTX_INPROC_SERVER)? };

        // Step 2: Query for IChatSettings
        let mut raw_ptr: *mut std::ffi::c_void = std::ptr::null_mut();
        unsafe {
            ole.query(&IID_IChatSettings, &mut raw_ptr).ok()?;
        }

        // Step 3: Cast to your bindgen interface
        let ptr = raw_ptr as *mut IChatSettings;
        Ok(Self::from_raw(ptr))
    }

    fn vtbl(&self) -> &IChatSettingsVtbl {
        unsafe { &*((*self.ptr).lpVtbl) }
    }

    // BackColor (OLE_COLOR → u32)
    pub fn get_back_color(&self) -> windows::core::Result<u32> {
        com_get!(self, get_BackColor, u32)
    }

    pub fn set_back_color(&self, val: Option<u32>) -> windows::core::Result<()> {
        com_put!(self, put_BackColor, val)
    }

    // ForeColor (OLE_COLOR → u32)
    pub fn get_fore_color(&self) -> windows::core::Result<u32> {
        com_get!(self, get_ForeColor, u32)
    }

    pub fn set_fore_color(&self, val: Option<u32>) -> windows::core::Result<()> {
        com_put!(self, put_ForeColor, val)
    }

    // RedirectURL (BSTR)
    pub fn get_redirect_url(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_RedirectURL)
    }

    pub fn set_redirect_url(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_RedirectURL, val)
    }

    // ResDLL (BSTR)
    pub fn get_res_dll(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_ResDLL)
    }

    pub fn set_res_dll(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_ResDLL, val)
    }
}
