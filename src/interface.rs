use com::{interfaces::IUnknown, sys::HRESULT};
use winapi::um::oaidl::{SAFEARRAY, VARIANT};
 
com::interfaces! {
    #[uuid("A43788C1-D91B-11D3-8F39-00C04F3651B8")]
    pub unsafe interface IRTDUpdateEvent: IUnknown {
        pub fn update_notify(&self) -> HRESULT;
        pub fn heartbeat_interval(&self, hb: *mut i32) -> HRESULT;
        pub fn disconnect(&self) -> HRESULT;
    }

    #[uuid("EC0E6191-DB51-11D3-8F3E-00C04F3651B8")]
    pub unsafe interface IRTDServer: IUnknown {
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
    }
}
