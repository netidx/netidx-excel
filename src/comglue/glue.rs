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
    mem,
    os::windows::ffi::{OsStrExt, OsStringExt},
    ptr,
    sync::mpsc,
    time::Duration,
};
use winapi::{
    shared::{
        guiddef::{GUID, IID_NULL},
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
        oaidl::{ITypeInfo, DISPID, DISPPARAMS, EXCEPINFO, SAFEARRAY, SAFEARRAYBOUND, VARIANT, VARIANT_n3},
        objidlbase::IStream,
        oleauto::{SafeArrayGetLBound, SafeArrayGetUBound, DISPATCH_METHOD, SafeArrayCreateVector, VariantInit, SysAllocStringLen},
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

unsafe fn irtd_update_event_loop(
    update_notify: DISPID,
    rx: mpsc::Receiver<()>,
    idp: *mut um::oaidl::IDispatch,
) {
    while let Ok(()) = rx.recv() {
        while let Ok(()) = rx.try_recv() {}
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
            &IID_NULL,
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
        let mut dispid = 0x0;
        debug!("get_dispids: calling GetIDsOfNames");
        let hr = (*idp).GetIDsOfNames(
            &IID_NULL,
            &mut update_notify.as_mut_ptr(),
            1,
            1000,
            &mut dispid,
        );
        debug!("update_event_thread: called GetIDsOfNames dispids: {:?}", dispid);
        if FAILED(hr) {
            error!("update_event_thread: could not get names {}", hr);
        }
        debug!("update_event_thread, init done, calling event loop");
        irtd_update_event_loop(dispid, args.rx, idp);
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

    pub(crate) fn update_notify(&self) {
        let _ = self.0.send(());
    }
}

struct VariantArray(*mut SAFEARRAY);

impl VariantArray {
    fn new(p: *mut SAFEARRAY) -> Self {
        VariantArray(p)
    }

    unsafe fn alloc(len: usize) -> Self {
        let p = SafeArrayCreateVector(wtypes::VT_VARIANT as u16, 0, len as u32);
        Self::new(p)
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

    unsafe fn typ_mut(&mut self) -> &mut u16 {
        &mut (*self.0).n1.n2_mut().vt
    }

    unsafe fn val(&self) -> &VARIANT_n3 {
        &(*self.0).n1.n2().n3
    }

    unsafe fn val_mut(&mut self) -> &mut VARIANT_n3 {
        &mut (*self.0).n1.n2_mut().n3
    }

    unsafe fn set_bool(&mut self, v: bool) {
        VariantInit(self.0);
        *self.typ_mut() = wtypes::VT_BOOL as u16;
        *self.val_mut().boolVal_mut() = if v { wtypes::VARIANT_TRUE } else { wtypes::VARIANT_FALSE };
    }

    unsafe fn set_null(&mut self) {
        VariantInit(self.0);
        *self.typ_mut() = wtypes::VT_NULL as u16;
    }

    unsafe fn set_error(&mut self) {
        VariantInit(self.0);
        *self.typ_mut() = wtypes::VT_ERROR as u16;
    }

    unsafe fn set_i32(&mut self, v: i32) {
        VariantInit(self.0);
        *self.typ_mut() = wtypes::VT_I4 as u16;
        *self.val_mut().lVal_mut() = v;
    }

    unsafe fn set_u32(&mut self, v: u32) {
        VariantInit(self.0);
        *self.typ_mut() = wtypes::VT_UI4 as u16;
        *self.val_mut().ulVal_mut() = v;
    }

    unsafe fn set_i64(&mut self, v: i64) {
        VariantInit(self.0);
        *self.typ_mut() = wtypes::VT_I8 as u16;
        *self.val_mut().llVal_mut() = v;
    }

    unsafe fn set_u64(&mut self, v: u64) {
        VariantInit(self.0);
        *self.typ_mut() = wtypes::VT_UI8 as u16;
        *self.val_mut().ullVal_mut() = v;
    }

    unsafe fn set_f32(&mut self, v: f32) {
        VariantInit(self.0);
        *self.typ_mut() = wtypes::VT_R4 as u16;
        *self.val_mut().fltVal_mut() = v;
    }

    unsafe fn set_f64(&mut self, v: f64) {
        VariantInit(self.0);
        *self.typ_mut() = wtypes::VT_R8 as u16;
        *self.val_mut().dblVal_mut() = v;
    }

    unsafe fn set_string(&mut self, v: &str) {
        VariantInit(self.0);
        *self.typ_mut() = wtypes::VT_BSTR as u16;
        let s = str_to_wstr(v);
        *self.val_mut().bstrVal_mut() = SysAllocStringLen(s.as_ptr(), s.len() as u32);
    }

    unsafe fn set_variant_array(&mut self, v: VariantArray) {
        VariantInit(self.0);
        *self.typ_mut() = (wtypes::VT_ARRAY | wtypes::VT_VARIANT) as u16;
        *self.val_mut().parray_mut() = v.0;
    }

    unsafe fn get_i32(&self) -> Result<i32> {
        if self.typ() == wtypes::VT_I4 as u16 {
            Ok(*self.val().lVal())
        } else {
            bail!("not a long value")
        }
    }

    unsafe fn get_byref_i32(&self) -> Result<*mut i32> {
        if self.typ() == (wtypes::VT_I4 | wtypes::VT_BYREF) as u16 {
            Ok(*self.val().plVal())
        } else {
            bail!("not a byref long value")
        }
    }

    unsafe fn get_path(&self) -> Result<Path> {
        if self.typ() == wtypes::VT_BSTR as u16 {
            let path = *self.val().bstrVal();
            let path = string_from_wstr(path);
            Ok(Path::from(ArcStr::from(&*path.to_string_lossy())))
        } else {
            bail!("not a string value")
        }
    }

    unsafe fn get_variant_array(&self) -> Result<VariantArray> {
        if self.typ() == (wtypes::VT_ARRAY | wtypes::VT_VARIANT) as u16 {
            Ok(VariantArray::new(*self.val().parray()))
        } else {
            bail!("not a variant array")
        }
    }

    unsafe fn get_irtd_update_event(&self) -> Result<IRTDUpdateEventWrap> {
        if self.typ() == wtypes::VT_DISPATCH as u16 {
            Ok(IRTDUpdateEventWrap::new(*self.val().pdispVal())?)
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
    let updates = params.get(0).get_irtd_update_event()?;
    server.server_start(updates);
    Ok(())
}

unsafe fn dispatch_connect_data(server: &Server, params: *mut DISPPARAMS) -> Result<()> {
    let params = Params::new(params)?;
    if params.len() != 3 {
        bail!("wrong number of args")
    }
    let topic_id = TopicId(params.get(2).get_i32()?);
    let topics = params.get(1).get_variant_array()?;
    if topics.len() == 0 {
        bail!("not enough topics")
    }
    let path = topics.get(0).get_path()?;
    Ok(server.connect_data(topic_id, path)?)
}

unsafe fn dispatch_refresh_data(server: &Server, params: *mut DISPPARAMS, result: &mut Variant) -> Result<()> {
    use netidx::subscriber::{Event, Value};
    let params = Params::new(params)?;
    debug!("param count: {}", params.len());
    debug!("param type: {}", params.get(0).typ());
    debug!("result type: {}", result.typ());
    let mut updates = server.refresh_data();
    let array = VariantArray::alloc(updates.len() * 2);
    for (i, (TopicId(tid), e)) in updates.drain().enumerate() {
        let i = i as isize;
        array.get(i).set_i32(tid);
        let i = i + 1;
        match e {
            Event::Unsubscribed => array.get(i).set_string("#SUB"),
            Event::Update(v) => match v {
                Value::I32(v) | Value::Z32(v) => array.get(i).set_i32(v),
                Value::U32(v) | Value::V32(v) => array.get(i).set_u32(v),
                Value::I64(v) | Value::Z64(v) => array.get(i).set_i64(v),
                Value::U64(v) | Value::V64(v) => array.get(i).set_u64(v),
                Value::F32(v) => array.get(i).set_f32(v),
                Value::F64(v) => array.get(i).set_f64(v),
                Value::True => array.get(i).set_bool(true),
                Value::False => array.get(i).set_bool(false),
                Value::String(s) => array.get(i).set_string(&*s),
                Value::Bytes(_) => array.get(i).set_string("#BIN"),
                Value::Null => array.get(i).set_null(),
                Value::Ok => array.get(i).set_string("OK"),
                Value::Error(e) => array.get(i).set_string(&format!("#ERR {}", &*e)),
                Value::Array(_) => array.get(i).set_string("#ARRAY"), // CR estokes: implement this?
                Value::DateTime(d) => array.get(i).set_string(&d.to_string()),
                Value::Duration(d) => array.get(i).set_string(&format!("{}s", d.as_secs_f64()))
            }
        }
    }
    result.set_variant_array(array);
    Ok(())
}

unsafe fn dispatch_disconnect_data(
    server: &Server,
    params: *mut DISPPARAMS,
) -> Result<()> {
    let params = Params::new(params)?;
    if params.len() != 1 {
        bail!("wrong number of args")
    }
    let topic_id = TopicId(params.get(0).get_i32()?);
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
            let mut result = Variant::new(result);
            match id {
                0 => {
                    debug!("ServerStart");
                    match dispatch_server_start(&self.server, params) {
                        Ok(()) => result.set_i32(1),
                        Err(e) => {
                            error!("server_start invalid arg {}", e);
                            result.set_error()
                        }
                    }
               },
                1 => {
                    debug!("ServerTerminate");
                    self.server.server_terminate();
                    result.set_null();
                },
                2 => {
                    debug!("ConnectData");
                    match dispatch_connect_data(&self.server, params) {
                        Ok(()) => result.set_i32(1),
                        Err(e) => {
                            error!("connect_data invalid arg {}", e);
                            result.set_error();
                        }
                    }
                },
                3 => {
                    debug!("RefreshData");
                    match dispatch_refresh_data(&self.server, params, &mut result) {
                        Ok(()) => (),
                        Err(e) => {
                            error!("refresh_data failed {}", e);
                            result.set_error()
                        }
                    }
                },
                4 => {
                    debug!("DisconnectData");
                    match dispatch_disconnect_data(&self.server, params) {
                        Ok(()) => result.set_i32(1),
                        Err(e) => {
                            error!("disconnect_data invalid arg {}", e);
                            result.set_error()
                        }
                    }
                },
                5 => {
                    debug!("Heartbeat");
                    result.set_i32(1);
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
