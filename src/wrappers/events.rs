use std::ffi::c_void;
use windows::{
    Win32::System::{
        Com::{DISPPARAMS, EXCEPINFO, IDispatch},
        Variant::VARIANT,
    },
    core::{GUID, HRESULT},
};

/// DISPID for OnRedirect from your IDL
const DISPID_ON_REDIRECT: i32 = 0x00000001;

/// A COM-compatible event sink for `_ICChatFrameEvents`
#[repr(C)]
pub struct ChatFrameEventSink {
    vtbl: *const IDispatchVtbl,
    ref_count: u32,
    pub on_redirect: Box<dyn Fn(String) + Send + Sync>,
}

#[repr(C)]
pub struct IDispatchVtbl {
    pub parent: [usize; 3], // IUnknown: QueryInterface, AddRef, Release
    pub GetTypeInfoCount: extern "system" fn(*mut c_void, *mut u32) -> HRESULT,
    pub GetTypeInfo: extern "system" fn(*mut c_void, u32, u32, *mut *mut c_void) -> HRESULT,
    pub GetIDsOfNames:
        extern "system" fn(*mut c_void, *const GUID, *mut *mut u16, u32, u32, *mut i32) -> HRESULT,
    pub Invoke: extern "system" fn(
        this: *mut c_void,
        dispid: i32,
        riid: *const GUID,
        lcid: u32,
        flags: u16,
        params: *mut DISPPARAMS,
        result: *mut VARIANT,
        excepinfo: *mut EXCEPINFO,
        arg_err: *mut u32,
    ) -> HRESULT,
}

// Static vtable instance
static VTABLE: IDispatchVtbl = IDispatchVtbl {
    parent: [0; 3], // You can implement IUnknown if needed
    GetTypeInfoCount: dummy_get_type_info_count,
    GetTypeInfo: dummy_get_type_info,
    GetIDsOfNames: dummy_get_ids_of_names,
    Invoke: invoke_handler,
};

// Dummy implementations for unused methods
extern "system" fn dummy_get_type_info_count(_: *mut c_void, _: *mut u32) -> HRESULT {
    HRESULT(0)
}
extern "system" fn dummy_get_type_info(
    _: *mut c_void,
    _: u32,
    _: u32,
    _: *mut *mut c_void,
) -> HRESULT {
    HRESULT(0)
}
extern "system" fn dummy_get_ids_of_names(
    _: *mut c_void,
    _: *const GUID,
    _: *mut *mut u16,
    _: u32,
    _: u32,
    _: *mut i32,
) -> HRESULT {
    HRESULT(0)
}

// Actual Invoke implementation
extern "system" fn invoke_handler(
    this: *mut c_void,
    dispid: i32,
    _riid: *const GUID,
    _lcid: u32,
    _flags: u16,
    params: *mut DISPPARAMS,
    _result: *mut VARIANT,
    _excepinfo: *mut EXCEPINFO,
    _arg_err: *mut u32,
) -> HRESULT {
    if dispid == DISPID_ON_REDIRECT {
        unsafe {
            let sink = &mut *(this as *mut ChatFrameEventSink);
            if (*params).cArgs == 1 {
                let arg = &*(*params).rgvarg;
                let raw_bstr = &arg.Anonymous.Anonymous.Anonymous.bstrVal;
                let string = raw_bstr.to_string();
                (sink.on_redirect)(string);
            }
        }
    }
    HRESULT(0)
}

impl ChatFrameEventSink {
    pub fn new<F: Fn(String) + Send + Sync + 'static>(handler: F) -> Box<Self> {
        Box::new(Self {
            vtbl: &VTABLE,
            ref_count: 1,
            on_redirect: Box::new(handler),
        })
    }

    pub fn as_idispatch(&mut self) -> *mut IDispatch {
        self as *mut _ as *mut IDispatch
    }
}
