use crate::interface::{IDispatch, IRTDServer, IRTDUpdateEvent};
use winapi::{
    shared::{minwindef::{WORD, UINT}, wtypesbase::LPOLESTR},
    um::{
        oaidl::{SAFEARRAY, VARIANT, ITypeInfo, DISPID, DISPPARAMS, EXCEPINFO}, 
        winnt::LCID,
        winbase::lstrlenW
    }, 
};
use com::{sys::{HRESULT, NOERROR, IID}};
use once_cell::sync::Lazy;
use log::{debug, LevelFilter};
use simplelog;
use std::{fs::File, os::windows::ffi::OsStringExt, ffi::OsString};

static LOGGER: Lazy<()> = Lazy::new(|| {
    let f = File::create("C:\\Users\\eric\\proj\\netidx-excel\\log.txt")
        .expect("couldn't open log file");
    simplelog::WriteLogger::init(LevelFilter::Debug, simplelog::Config::default(), f)
        .expect("couldn't init log")
});

fn maybe_init_logger() {
    *LOGGER
}

com::class! {
    #[derive(Debug)]
    pub class NetidxRTD: IRTDServer(IDispatch) {}

    impl IDispatch for NetidxRTD {
        fn get_type_info_count(&self, info: *mut UINT) -> HRESULT { 
            maybe_init_logger();
            debug!("get_type_info_count(info: {})", unsafe { *info });
            unsafe { *info = 0; } // no we don't support type info
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
            let names = unsafe { std::slice::from_raw_parts(names, names_len as usize) };
            for name in names {
                let name = unsafe { std::slice::from_raw_parts(*name, lstrlenW(*name) as usize) };
                let s = OsString::from_wide(name);
                match s.into_string() {
                    Err(_) => debug!("excel sent us invalid unicode"),
                    Ok(s) => debug!("name: {}", s)
                }
            }
            NOERROR
        }

        fn invoke(
            &self, 
            _id: DISPID, 
            _iid: *const IID, 
            _lcid: LCID, 
            _flags: WORD, 
            _params: *mut DISPPARAMS,
            _result: *mut VARIANT,
            _exception: *mut EXCEPINFO,
            _arg_error: *mut UINT
        ) -> HRESULT { NOERROR }
    }

    impl IRTDServer for NetidxRTD {
        fn server_start(&self, _cb: *const IRTDUpdateEvent, _res: *mut i32) -> HRESULT {
            std::fs::write("C:\\Users\\eric\\proj\\netidx-excel\\log.txt", "I was initialized").unwrap();
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
    }
}
