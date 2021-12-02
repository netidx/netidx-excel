use crate::interface::{IDispatch, IRTDServer, IRTDUpdateEvent, IID_IRTDSERVER, IID_IRTDUPDATE_EVENT};
use winapi::{
    shared::{
        guiddef::GUID,
        minwindef::{WORD, UINT}, 
        wtypesbase::LPOLESTR,
        wtypes::{
            VT_DISPATCH,
        }
    },
    um::{
        self,
        oaidl::{SAFEARRAY, VARIANT, ITypeInfo, DISPID, DISPPARAMS, EXCEPINFO},
        oleauto::DISPATCH_METHOD,
        winnt::LCID,
        winbase::lstrlenW
    }, 
};
use oaidl::{VariantExt, SafeArrayExt};
use com::{sys::{HRESULT, NOERROR, IID}};
use once_cell::sync::Lazy;
use log::{debug, LevelFilter};
use simplelog;
use std::{fs::File, os::windows::ffi::{OsStringExt, OsStrExt}, ffi::OsString, borrow::Cow};

static LOGGER: Lazy<()> = Lazy::new(|| {
    let f = File::create("C:\\Users\\eric\\proj\\netidx-excel\\log.txt")
        .expect("couldn't open log file");
    simplelog::WriteLogger::init(LevelFilter::Debug, simplelog::Config::default(), f)
        .expect("couldn't init log")
});

fn maybe_init_logger() {
    *LOGGER
}

unsafe fn string_from_wstr<'a>(s: *const u16) -> OsString {
    OsString::from_wide(std::slice::from_raw_parts(s, lstrlenW(s) as usize))
}

fn str_to_wstr(s: &str) -> Vec<u16> {
    let mut v = OsString::from(s).encode_wide().collect::<Vec<_>>();
    v.push(0);
    v
}


// Excel hands us an IRTDUpdateEvent class that we need to use to tell it that we have data, 
// however it doesn't give us an IRTDUpdateEvent COM interface, it gives us an IDispatch COM
// interface, so we need to use that to call the methods of IRTDUpdateEvent through IDispatch.
struct IRTDUpdateEventWrap {
    ptr: *mut um::oaidl::IDispatch,
    iid: GUID,
    update_notify_id: DISPID,
    heartbeat_interval_id: DISPID,
    disconnect_id: DISPID,
}

impl IRTDUpdateEventWrap {
    fn new(ptr: *mut um::oaidl::IDispatch) -> Self {
        assert!(!ptr.is_null());
        let mut update_notify = str_to_wstr("UpdateNotify");
        let mut heartbeat_interval = str_to_wstr("HeartbeatInterval");
        let mut disconnect = str_to_wstr("Disconnect");
        let mut names = [update_notify.as_mut_ptr(), heartbeat_interval.as_mut_ptr(), disconnect.as_mut_ptr()];
        let mut dispids: [DISPID; 3] = [0x0, 0x0, 0x0];
        let iid = GUID {
            Data1: IID_IRTDUPDATE_EVENT.data1,
            Data2: IID_IRTDUPDATE_EVENT.data2,
            Data3: IID_IRTDUPDATE_EVENT.data3,
            Data4: IID_IRTDUPDATE_EVENT.data4,
        };
        let res = unsafe {
            (*ptr).GetIDsOfNames(&iid, names.as_mut_ptr(), 3, 0, dispids.as_mut_ptr())
        };
        if res != NOERROR {
            panic!("IRTDUpdateEventWrap: could not get names {}", res);
        }
        IRTDUpdateEventWrap {
            ptr,
            iid,
            update_notify_id: dispids[0],
            heartbeat_interval_id: dispids[1],
            disconnect_id: dispids[2],
        }
    }

    fn update_notify(&self) {
        
    }
}

com::class! {
    #[derive(Debug)]
    pub class NetidxRTD: IRTDServer(IDispatch) {}

    impl IDispatch for NetidxRTD {
        fn get_type_info_count(&self, info: *mut UINT) -> HRESULT { 
            maybe_init_logger();
            debug!("get_type_info_count(info: {})", unsafe { *info });
            if !info.is_null() {
                unsafe { *info = 0; } // no we don't support type info
            }
            NOERROR 
        }
        fn get_type_info(&self, _lcid: LCID, _type_info: *mut *mut ITypeInfo) -> HRESULT { NOERROR }

        pub fn get_ids_of_names(
            &self, 
            riid: *const IID, 
            names: *const LPOLESTR, 
            names_len: UINT, 
            lcid: LCID, 
            ids: *mut DISPID
        ) -> HRESULT {
            maybe_init_logger();
            debug!("get_ids_of_names(riid: {:?}, names: {:?}, names_len: {}, lcid: {}, ids: {:?})", riid, names, names_len, lcid, ids);
            if !ids.is_null() && !names.is_null() && !riid.is_null() && unsafe { (*riid) == IID_IRTDSERVER } {
                for i in 0..names_len {
                    let name = unsafe { string_from_wstr(*names.offset(i as isize)) };
                    let name = name.to_string_lossy();
                    debug!("name: {}", name);
                    match &*name {
                        "ServerStart" => unsafe { *ids.offset(i as isize) = 0; },
                        "ServerTerminate" => unsafe { *ids.offset(i as isize) = 1; }
                        "ConnectData" => unsafe { *ids.offset(i as isize) = 2; }
                        "RefreshData" => unsafe { *ids.offset(i as isize) = 3; }
                        "DisconnectData" => unsafe { *ids.offset(i as isize) = 4; }
                        "Heartbeat" => unsafe { *ids.offset(i as isize) = 5; }
                        _ => debug!("unknown method: {}", name)
                    }
                }
            }
            NOERROR
        }

        fn invoke(
            &self, 
            id: DISPID, 
            iid: *const IID, 
            lcid: LCID, 
            flags: WORD, 
            params: *mut DISPPARAMS,
            result: *mut VARIANT,
            exception: *mut EXCEPINFO,
            arg_error: *mut UINT
        ) -> HRESULT { 
            maybe_init_logger();
            debug!(
                "invoke(id: {}, iid: {:?}, lcid: {}, flags: {}, params: {:?}, result: {:?}, exception: {:?}, arg_error: {:?})", 
                id, iid, lcid, flags, params, result, exception, arg_error
            );
            unsafe {
                for i in 0..(*params).cArgs as usize {
                    let arg = (*params).rgvarg.offset(i as isize);
                    debug!("arg type: {}", (*arg).n1.n2().vt);
                }
            }
            NOERROR 
        }
    }

    impl IRTDServer for NetidxRTD {
        fn server_start(&self, _cb: *const IRTDUpdateEvent, _res: *mut i32) -> HRESULT {
            NOERROR
        }

        fn connect_data(&self, _topic_id: i32, _topic: *const SAFEARRAY, _get_new_values: *mut VARIANT, _res: *mut VARIANT) -> HRESULT {
            NOERROR
        }

        fn refresh_data(&self, _topic_count: *mut i32, _data: *mut SAFEARRAY) -> HRESULT {
            NOERROR
        }

        fn disconnect_data(&self, _topic_id: i32) -> HRESULT {
            NOERROR
        }

        fn heartbeat(&self, _res: *mut i32) -> HRESULT {
            NOERROR
        }

        fn server_terminate(&self) -> HRESULT {
            NOERROR
        }
    }
}
