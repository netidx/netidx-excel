use crate::interface::{IDispatch, IRTDServer, IRTDUpdateEvent};
use winapi::{
    shared::{minwindef::{WORD, UINT}, wtypesbase::LPOLESTR},
    um::{
        oaidl::{SAFEARRAY, VARIANT, ITypeInfo, DISPID, DISPPARAMS, EXCEPINFO}, 
        winnt::LCID
    }, 
};
use com::{sys::{HRESULT, NOERROR, IID}};
use once_cell::sync::Lazy;
use log::{debug, LevelFilter};
use simplelog;
use std::fs::File;

static LOGGER: Lazy<()> = Lazy::new(|| {
    let f = File::create("C:\\Users\\eric\\proj\\netidx-excel\\log.txt")
        .expect("couldn't open log file");
    simplelog::WriteLogger::init(LevelFilter::Debug, simplelog::Config::default(), f)
        .expect("couldn't init log")
});

com::class! {
    #[derive(Debug)]
    pub class NetidxRTD: IRTDServer(IDispatch) {}

    impl IDispatch for NetidxRTD {
        fn get_type_info_count(&self, info: *mut UINT) -> HRESULT { 
            *LOGGER;
            debug!("get_type_info_count: {}", unsafe { *info });
            unsafe { *info = 0; } // no we don't support type info
            NOERROR 
        }
        fn get_type_info(&self, _lcid: LCID, _type_info: *mut *mut ITypeInfo) -> HRESULT { NOERROR }

        pub fn get_ids_of_names(
            &self, 
            _riid: *const IID, 
            _names: *const LPOLESTR, 
            _names_len: UINT, 
            _lcid: LCID, 
            _ids: *mut DISPID
        ) -> HRESULT {
            std::fs::write("C:\\Users\\eric\\proj\\netidx-excel\\log.txt", "ids of names called").unwrap();
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
