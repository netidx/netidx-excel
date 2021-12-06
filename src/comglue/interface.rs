use com::{
    interfaces::IUnknown,
    sys::{HRESULT, IID},
};
use windows::Win32::System::Com::{ITypeInfo, DISPPARAMS, EXCEPINFO, SAFEARRAY, VARIANT};

// bde5f32a-14d9-414e-a0af-8390a1601944
pub const CLSID: IID = IID {
    data1: 0xbde5f32a,
    data2: 0x14d9,
    data3: 0x414e,
    data4: [0xa0, 0xaf, 0x83, 0x90, 0xa1, 0x60, 0x19, 0x44],
};

com::interfaces! {
    #[uuid("00020400-0000-0000-C000-000000000046")]
    pub unsafe interface IDispatch: IUnknown {
        pub fn get_type_info_count(&self, info: *mut u32) -> HRESULT;
        pub fn get_type_info(&self, lcid: u32, type_info: *mut *mut ITypeInfo) -> HRESULT;
        pub fn get_ids_of_names(
            &self,
            riid: *const IID,
            names: *const *mut u16,
            names_len: u32,
            lcid: u32,
            ids: *mut i32
        ) -> HRESULT;
        pub fn invoke(
            &self,
            id: i32,
            iid: *const IID,
            lcid: u32,
            flags: u16,
            params: *mut DISPPARAMS,
            result: *mut VARIANT,
            exception: *mut EXCEPINFO,
            arg_error: *mut u32
        ) -> HRESULT;
    }

    #[uuid("A43788C1-D91B-11D3-8F39-00C04F3651B8")]
    pub unsafe interface IRTDUpdateEvent: IDispatch {
        pub fn update_notify(&self) -> HRESULT;
        pub fn heartbeat_interval(&self, hb: *mut i32) -> HRESULT;
        pub fn disconnect(&self) -> HRESULT;
    }

    #[uuid("EC0E6191-DB51-11D3-8F3E-00C04F3651B8")]
    pub unsafe interface IRTDServer: IDispatch {
        pub fn server_start(&self, cb: *const IRTDUpdateEvent, res: *mut i32) -> HRESULT;
        pub fn connect_data(
            &self,
            topic_id: i32,
            topic: *const SAFEARRAY,
            get_new_values: *mut VARIANT,
            res: *mut VARIANT
        ) -> HRESULT;
        pub fn refresh_data(&self, topic_count: *mut i32, data: *mut SAFEARRAY) -> HRESULT;
        pub fn disconnect_data(&self, topic_id: i32) -> HRESULT;
        pub fn heartbeat(&self, res: *mut i32) -> HRESULT;
        pub fn server_terminate(&self) -> HRESULT;
    }
}
