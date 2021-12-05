use crate::{
    comglue::interface::{
        IDispatch, IRTDServer, IRTDUpdateEvent, IID_IDISPATCH, IID_IRTDSERVER,
        IID_IRTDUPDATE_EVENT,
    },
    server::{Server, TopicId},
};
use anyhow::{bail, Result};
use arcstr::ArcStr;
use com::{
    interfaces::IUnknown,
    sys::{HRESULT, IID, NOERROR},
};
use log::{debug, error, LevelFilter};
use netidx::path::Path;
use oaidl::{SafeArrayExt, VariantExt, VtNull};
use once_cell::sync::Lazy;
use simplelog;
use std::{
    boxed::Box,
    ffi::{c_void, OsString},
    fs::File,
    marker::{Send, Sync},
    os::windows::ffi::{OsStrExt, OsStringExt},
    ptr,
    sync::mpsc,
    mem
};
use winapi::{
    shared::{
        guiddef::GUID,
        minwindef::{UINT, WORD},
        winerror::{ERROR_CREATE_FAILED, FAILED, SUCCEEDED},
        wtypes,
        wtypesbase::LPOLESTR,
    },
    um::{
        self,
        combaseapi::{
            CoGetInterfaceAndReleaseStream, CoInitializeEx,
            CoMarshalInterThreadInterfaceInStream, CoUninitialize,
        },
        oaidl::{ITypeInfo, DISPID, DISPPARAMS, EXCEPINFO, SAFEARRAY, VARIANT},
        objidlbase::IStream,
        oleauto::{SafeArrayGetLBound, SafeArrayGetUBound, DISPATCH_METHOD},
        processthreadsapi::CreateThread,
        winbase::lstrlenW,
        winnt::LCID,
    },
};

static LOGGER: Lazy<()> = Lazy::new(|| {
    let f = File::create("C:\\Users\\eric\\proj\\netidx-excel\\log.txt")
        .expect("couldn't open log file");
    simplelog::WriteLogger::init(LevelFilter::Debug, simplelog::Config::default(), f)
        .expect("couldn't init log")
});

pub fn maybe_init_logger() {
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

unsafe fn set_variant_lval(v: *mut VARIANT, l: i32) {
    *(*v).n1.n2_mut().n3.lVal_mut() = l;
}

// IRTDUpdateEvent is single apartment threaded, and that means we need to ask COM
// to make a proxy for us in order to run it in another thread. Since we MUST run in
// another thread to be async, this is mandatory. We have to marshal the interface when
// we receive it, and then unmarshal it in the update thread, which is then able to
// call into it.
struct IRTDUpdateEventThreadArgs {
    stream: *mut IStream,
    rx: mpsc::Receiver<()>,
}

static IDISPATCH_GUID: GUID = GUID {
    Data1: IID_IDISPATCH.data1,
    Data2: IID_IDISPATCH.data2,
    Data3: IID_IDISPATCH.data3,
    Data4: IID_IDISPATCH.data4,
};

static IRTDUPDATE_EVENT_GUID: GUID = GUID {
    Data1: IID_IRTDUPDATE_EVENT.data1,
    Data2: IID_IRTDUPDATE_EVENT.data2,
    Data3: IID_IRTDUPDATE_EVENT.data3,
    Data4: IID_IRTDUPDATE_EVENT.data4,
};

unsafe fn irtd_update_event_loop(
    update_notify: DISPID,
    rx: mpsc::Receiver<()>,
    idp: *mut um::oaidl::IDispatch,
) {
    while let Ok(()) = rx.recv() {
        let mut args = [];
        let mut named_args = [];
        let mut params = DISPPARAMS {
            rgvarg: args.as_mut_ptr(),
            rgdispidNamedArgs: named_args.as_mut_ptr(),
            cArgs: 0,
            cNamedArgs: 0,
        };
        let mut _result =
            VariantExt::into_variant(VtNull).expect("couldn't create result variant");
        let mut _arg_err = 0;
        let hr = (*idp).Invoke(
            update_notify,
            &IRTDUPDATE_EVENT_GUID,
            0,
            DISPATCH_METHOD,
            &mut params,
            _result.as_ptr(),
            ptr::null_mut(),
            &mut _arg_err,
        );
        if FAILED(hr) {
            error!("IRTDUpdateEvent: update_notify failed {}", hr);
        }
    }
}

unsafe extern "system" fn irtd_update_event_thread(ptr: *mut c_void) -> u32 {
    maybe_init_logger();
    let args = Box::from_raw(ptr.cast::<IRTDUpdateEventThreadArgs>());
    let hr = CoInitializeEx(ptr::null_mut(), 0);
    if FAILED(hr) {
        error!("update_event_thread: failed to initialize COM {}", hr)
    }
    let mut idp: *mut um::oaidl::IDispatch = ptr::null_mut();
    let hr = CoGetInterfaceAndReleaseStream(
        args.stream,
        &IDISPATCH_GUID,
        ((&mut idp) as *mut *mut um::oaidl::IDispatch).cast::<*mut c_void>(),
    );
    if FAILED(hr) {
        error!("update_event_thread: failed to unmarshal the IDispatch interface {}", hr);
    }
    if !idp.is_null() {
        let mut update_notify = str_to_wstr("UpdateNotify");
        let mut heartbeat_interval = str_to_wstr("HeartbeatInterval");
        let mut disconnect = str_to_wstr("Disconnect");
        let mut names = [
            update_notify.as_mut_ptr(),
            heartbeat_interval.as_mut_ptr(),
            disconnect.as_mut_ptr(),
        ];
        let mut dispids: [DISPID; 3] = [0x0, 0x0, 0x0];
        debug!("get_dispids: calling GetIDsOfNames");
        let hr = unsafe {
            (*idp).GetIDsOfNames(
                &IRTDUPDATE_EVENT_GUID,
                names.as_mut_ptr(),
                3,
                0,
                dispids.as_mut_ptr(),
            )
        };
        debug!("update_event_thread: called GetIDsOfNames dispids: {:?}", dispids);
        if FAILED(hr) {
            error!("update_event_thread: could not get names {}", hr);
        }
        debug!("update_event_thread, init done, calling event loop");
        irtd_update_event_loop(dispids[0], args.rx, idp);
    }
    CoUninitialize();
    0
}

pub(crate) struct IRTDUpdateEventWrap(mpsc::Sender<()>);

impl IRTDUpdateEventWrap {
    unsafe fn new(ptr: *mut um::oaidl::IDispatch) -> Result<Self> {
        use winapi::um::unknwnbase::IUnknown;
        assert!(!ptr.is_null());
        let (tx, rx) = mpsc::channel();
        let mut args =
            Box::new(IRTDUpdateEventThreadArgs { stream: ptr::null_mut(), rx });
        let res = CoMarshalInterThreadInterfaceInStream(
            &IDISPATCH_GUID,
            mem::transmute::<&IUnknown, *mut IUnknown>(&*ptr),
            &mut args.stream,
        );
        if FAILED(res) {
            bail!("couldn't marshal interface {}", res);
        }
        let mut threadid = 0u32;
        CreateThread(
            ptr::null_mut(),
            0,
            Some(irtd_update_event_thread),
            Box::into_raw(args).cast::<c_void>(),
            0,
            &mut threadid,
        );
        Ok(IRTDUpdateEventWrap(tx))
    }

    pub(crate) fn update_notify(&mut self) {
        let _ = self.0.send(());
    }
}

struct VariantArray(*mut SAFEARRAY);

impl VariantArray {
    fn new(p: *mut SAFEARRAY) -> Self {
        VariantArray(p)
    }

    unsafe fn len(&self) -> usize {
        let mut lbound = 0;
        let mut ubound = 0;
        SafeArrayGetLBound(self.0, 1, &mut lbound);
        SafeArrayGetUBound(self.0, 1, &mut ubound);
        (1 + ubound - lbound) as usize
    }

    unsafe fn get(&self, i: isize) -> Variant {
        Variant::new((*self.0).pvData.cast::<VARIANT>().offset(i))
    }
}

struct Variant(*mut VARIANT);

impl Variant {
    fn new(p: *mut VARIANT) -> Self {
        Variant(p)
    }

    unsafe fn typ(&self) -> u16 {
        (*self.0).n1.n2().vt
    }

    unsafe fn as_long(&self) -> Result<i32> {
        if self.typ() == wtypes::VT_I4 as u16 {
            Ok(*(*self.0).n1.n2().n3.lVal())
        } else {
            bail!("not a long value")
        }
    }

    unsafe fn as_path(&self) -> Result<Path> {
        if self.typ() == wtypes::VT_BSTR as u16 {
            let path = *(*self.0).n1.n2().n3.bstrVal();
            let path = string_from_wstr(path);
            Ok(Path::from(ArcStr::from(&*path.to_string_lossy())))
        } else {
            bail!("not a string value")
        }
    }

    unsafe fn as_variant_array(&self) -> Result<VariantArray> {
        if self.typ() == (wtypes::VT_ARRAY | wtypes::VT_VARIANT) as u16 {
            Ok(VariantArray::new(*(*self.0).n1.n2().n3.parray()))
        } else {
            bail!("not a variant array")
        }
    }

    unsafe fn as_irtd_update_event(&self) -> Result<IRTDUpdateEventWrap> {
        if self.typ() == wtypes::VT_DISPATCH as u16 {
            Ok(IRTDUpdateEventWrap::new(*(*self.0).n1.n2().n3.pdispVal())?)
        } else {
            bail!("not an update event interface")
        }
    }
}

struct Params(*mut DISPPARAMS);

impl Params {
    fn new(ptr: *mut DISPPARAMS) -> Result<Self> {
        if ptr.is_null() {
            bail!("invalid params")
        }
        Ok(Params(ptr))
    }

    unsafe fn len(&self) -> usize {
        (*self.0).cArgs as usize
    }

    unsafe fn get(&self, i: isize) -> Variant {
        Variant::new((*self.0).rgvarg.offset(i))
    }
}

unsafe fn dispatch_server_start(server: &Server, params: *mut DISPPARAMS) -> Result<()> {
    let params = Params::new(params)?;
    let updates = params.get(0).as_irtd_update_event()?;
    server.server_start(updates);
    Ok(())
}

unsafe fn dispatch_connect_data(server: &Server, params: *mut DISPPARAMS) -> Result<()> {
    let params = Params::new(params)?;
    if params.len() != 3 {
        bail!("wrong number of args")
    }
    let topic_id = TopicId(params.get(2).as_long()?);
    let topics = params.get(1).as_variant_array()?;
    if topics.len() == 0 {
        bail!("not enough topics")
    }
    let path = topics.get(0).as_path()?;
    Ok(server.connect_data(topic_id, path)?)
}

unsafe fn dispatch_disconnect_data(
    server: &Server,
    params: *mut DISPPARAMS,
) -> Result<()> {
    let params = Params::new(params)?;
    if params.len() != 1 {
        bail!("wrong number of args")
    }
    let topic_id = TopicId(params.get(0).as_long()?);
    Ok(server.disconnect_data(topic_id))
}

com::class! {
    #[derive(Debug)]
    pub class NetidxRTD: IRTDServer(IDispatch) {
        server: Server,
    }

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
            if !ids.is_null() && !names.is_null() {
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

        unsafe fn invoke(
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
            assert!(!params.is_null());
            match id {
                0 => {
                    debug!("ServerStart");
                    match dispatch_server_start(&self.server, params) {
                        Ok(()) => set_variant_lval(result, 1),
                        Err(e) => {
                            error!("server_start invalid arg {}", e);
                            set_variant_lval(result, 0)
                        }
                    }
               },
                1 => {
                    debug!("ServerTerminate");
                    self.server.server_terminate();
                    set_variant_lval(result, 1);
                },
                2 => {
                    debug!("ConnectData");
                    match dispatch_connect_data(&self.server, params) {
                        Ok(()) => set_variant_lval(result, 1),
                        Err(e) => {
                            error!("connect_data invalid arg {}", e);
                            set_variant_lval(result, 0)
                        }
                    }
                },
                3 => {
                    debug!("RefreshData")
                },
                4 => {
                    debug!("DisconnectData");
                    match dispatch_disconnect_data(&self.server, params) {
                        Ok(()) => set_variant_lval(result, 1),
                        Err(e) => {
                            error!("disconnect_data invalid arg {}", e);
                            set_variant_lval(result, 0)
                        }
                    }
                },
                5 => {
                    debug!("Heartbeat");
                    set_variant_lval(result, 1);
                },
                _ => {
                    debug!("unknown method {} called", id)
                },
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
