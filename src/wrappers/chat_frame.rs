use windows::{
    core::{Interface, Result, GUID}, Win32::System::{
        Com::{CoCreateInstance, CLSCTX_INPROC_SERVER},
        Ole::{IOleObject},
    }
};

use crate::bindings::{
    guids::{self, CLSID_MSNChatFrame, IID_IChatFrame},
    ichat_frame::{IChatFrame, IChatFrameVtbl},
};

#[derive(Clone)]
pub struct ChatFrame {
    ptr: *mut IChatFrame,
}

unsafe impl Interface for ChatFrame {
    type Vtable = IChatFrameVtbl; // from bindgen

    const IID: GUID = IID_IChatFrame;

    fn as_raw(&self) -> *mut std::ffi::c_void {
        self.ptr as *mut _
    }

    unsafe fn from_raw(raw: *mut std::ffi::c_void) -> Self {
        Self { ptr: raw as *mut IChatFrame }
    }

    fn vtable(&self) -> &Self::Vtable {
        unsafe {
            &*(*self.ptr).lpVtbl
        }
    }
}

impl ChatFrame {
    /// Constructs a `ChatFrame` from a raw COM pointer.
    pub unsafe fn from_raw(ptr: *mut IChatFrame) -> Self {
        Self { ptr }
    }

    pub fn as_ptr(&self) -> *mut IChatFrame {
        self.ptr
    }

    /// Creates a new `ChatFrame` instance via `CoCreateInstance`.
    pub fn create() -> Result<Self> {
        // Step 1: Create the object using a known interface
        let ole: IOleObject =
            unsafe { CoCreateInstance(&CLSID_MSNChatFrame, None, CLSCTX_INPROC_SERVER)? };

        // Step 2: Query for IChatFrame
        let mut raw_ptr: *mut std::ffi::c_void = std::ptr::null_mut();
        unsafe {
            ole.query(&IID_IChatFrame, &mut raw_ptr).ok()?;
        }

        // Step 3: Cast to your bindgen interface
        let ptr = raw_ptr as *mut IChatFrame;
        Ok(unsafe { Self::from_raw(ptr) })
    }

    fn vtbl(&self) -> &IChatFrameVtbl {
        unsafe { &*((*self.ptr).lpVtbl) }
    }

    // OLE_COLOR properties (u32)
    pub fn get_back_color(&self) -> windows::core::Result<u32> {
        com_get!(self, get_BackColor, u32)
    }
    pub fn set_back_color(&self, val: Option<u32>) -> windows::core::Result<()> {
        com_put!(self, put_BackColor, val)
    }

    pub fn get_back_highlight_color(&self) -> windows::core::Result<u32> {
        com_get!(self, get_BackHighlightColor, u32)
    }
    pub fn set_back_highlight_color(&self, val: Option<u32>) -> windows::core::Result<()> {
        com_put!(self, put_BackHighlightColor, val)
    }

    pub fn get_button_frame_color(&self) -> windows::core::Result<u32> {
        com_get!(self, get_ButtonFrameColor, u32)
    }
    pub fn set_button_frame_color(&self, val: Option<u32>) -> windows::core::Result<()> {
        com_put!(self, put_ButtonFrameColor, val)
    }

    pub fn get_top_back_highlight_color(&self) -> windows::core::Result<u32> {
        com_get!(self, get_TopBackHighlightColor, u32)
    }
    pub fn set_top_back_highlight_color(&self, val: Option<u32>) -> windows::core::Result<()> {
        com_put!(self, put_TopBackHighlightColor, val)
    }

    pub fn get_input_border_color(&self) -> windows::core::Result<u32> {
        com_get!(self, get_InputBorderColor, u32)
    }
    pub fn set_input_border_color(&self, val: Option<u32>) -> windows::core::Result<()> {
        com_put!(self, put_InputBorderColor, val)
    }

    pub fn get_button_text_color(&self) -> windows::core::Result<u32> {
        com_get!(self, get_ButtonTextColor, u32)
    }
    pub fn set_button_text_color(&self, val: Option<u32>) -> windows::core::Result<()> {
        com_put!(self, put_ButtonTextColor, val)
    }

    pub fn get_button_back_color(&self) -> windows::core::Result<u32> {
        com_get!(self, get_ButtonBackColor, u32)
    }
    pub fn set_button_back_color(&self, val: Option<u32>) -> windows::core::Result<()> {
        com_put!(self, put_ButtonBackColor, val)
    }

    pub fn get_chat_mode(&self) -> windows::core::Result<i32> {
        com_get!(self, get_ChatMode, i32)
    }
    pub fn set_chat_mode(&self, val: Option<i32>) -> windows::core::Result<()> {
        com_put!(self, put_ChatMode, val)
    }

    pub fn get_feature(&self) -> windows::core::Result<u32> {
        com_get!(self, get_Feature, u32)
    }
    pub fn set_feature(&self, val: Option<u32>) -> windows::core::Result<()> {
        com_put!(self, put_Feature, val)
    }

    // BSTR properties
    pub fn get_room_name(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_RoomName)
    }
    pub fn set_room_name(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_RoomName, val)
    }

    pub fn get_hex_room_name(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_HexRoomName)
    }
    pub fn set_hex_room_name(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_HexRoomName, val)
    }

    pub fn get_nick_name(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_NickName)
    }
    pub fn set_nick_name(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_NickName, val)
    }

    pub fn get_server(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_Server)
    }
    pub fn set_server(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_Server, val)
    }

    pub fn get_url_back(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_URLBack)
    }
    pub fn set_url_back(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_URLBack, val)
    }

    pub fn get_category(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_Category)
    }
    pub fn set_category(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_Category, val)
    }

    pub fn get_topic(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_Topic)
    }
    pub fn set_topic(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_Topic, val)
    }

    pub fn get_welcome_msg(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_WelcomeMsg)
    }
    pub fn set_welcome_msg(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_WelcomeMsg, val)
    }

    pub fn get_base_url(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_BaseURL)
    }
    pub fn set_base_url(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_BaseURL, val)
    }

    pub fn get_create_room(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_CreateRoom)
    }
    pub fn set_create_room(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_CreateRoom, val)
    }

    pub fn get_chat_home(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_ChatHome)
    }
    pub fn set_chat_home(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_ChatHome, val)
    }

    pub fn get_locale(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_Locale)
    }
    pub fn set_locale(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_Locale, val)
    }

    pub fn get_res_dll(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_ResDLL)
    }
    pub fn set_res_dll(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_ResDLL, val)
    }

    pub fn get_passport_ticket(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_PassportTicket)
    }
    pub fn set_passport_ticket(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_PassportTicket, val)
    }

    pub fn get_passport_profile(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_PassportProfile)
    }
    pub fn set_passport_profile(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_PassportProfile, val)
    }

    pub fn get_message_of_the_day(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_MessageOfTheDay)
    }
    pub fn set_message_of_the_day(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_MessageOfTheDay, val)
    }

    pub fn get_channel_language(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_ChannelLanguage)
    }
    pub fn set_channel_language(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_ChannelLanguage, val)
    }

    pub fn get_invitation_code(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_InvitationCode)
    }

    pub fn set_invitation_code(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_InvitationCode, val)
    }

    pub fn get_nickname_to_invite(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_NicknameToInvite)
    }
    pub fn set_nickname_to_invite(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_NicknameToInvite, val)
    }

    pub fn get_msnreg_cookie(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_MSNREGCookie)
    }
    pub fn set_msnreg_cookie(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_MSNREGCookie, val)
    }

    pub fn get_creation_modes(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_CreationModes)
    }
    pub fn set_creation_modes(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_CreationModes, val)
    }

    pub fn get_msn_profile(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_MSNProfile)
    }
    pub fn set_msn_profile(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_MSNProfile, val)
    }

    pub fn get_market(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_Market)
    }
    pub fn set_market(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_Market, val)
    }

    pub fn get_whisper_content(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_WhisperContent)
    }
    pub fn set_whisper_content(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_WhisperContent, val)
    }

    pub fn get_user_role(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_UserRole)
    }
    pub fn set_user_role(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_UserRole, val)
    }

    pub fn get_audit_message(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_AuditMessage)
    }
    pub fn set_audit_message(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_AuditMessage, val)
    }

    pub fn get_subscriber_info(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_SubscriberInfo)
    }
    pub fn set_subscriber_info(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_SubscriberInfo, val)
    }

    pub fn get_upsell_url(&self) -> windows::core::Result<String> {
        com_get_bstr!(self, get_UpsellURL)
    }
    pub fn set_upsell_url(&self, val: Option<&str>) -> windows::core::Result<()> {
        com_put_bstr!(self, put_UpsellURL, val)
    }
}
