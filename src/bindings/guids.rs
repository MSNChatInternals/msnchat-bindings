use windows::core::GUID;

// 📚 Type Library
pub const LIBID_MSNChat: GUID = GUID::from_u128(0x0f0a655c_6c6d_4e0b_8038_f980b36f9c78);

// 🧩 Interface IIDs
pub const IID_IChatFrame: GUID = GUID::from_u128(0x125e64fa_3304_4bb9_a756_d0d44cc8cd7d);
pub const IID_IChatSettings: GUID = GUID::from_u128(0xd5ef4299_12f1_474d_98c5_3c658fd2e343);
pub const IID_ICChatFrameEvents: GUID = GUID::from_u128(0x5eeb8014_53b2_448b_9f3b_c553424832e1);

// 🧱 CoClass CLSIDs
pub const CLSID_MSNChatFrame: GUID = GUID::from_u128(0xf58e1cef_a068_4d2f_b3c2_9a5be42525f8);
pub const CLSID_ChatSettings: GUID = GUID::from_u128(0xfa980e7e_9e44_4d2f_b3c2_9a5be42525f8);
