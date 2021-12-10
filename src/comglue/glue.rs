use crate::{
    comglue::{
        interface::{IDispatch, IRTDServer, IRTDUpdateEvent, IID_IDISPATCH},
        variant::{str_to_wstr, string_from_wstr, SafeArray, Variant},
    },
    server::{Server, TopicId},
};
use anyhow::{anyhow, bail, Result};
use arcstr::ArcStr;
use com::sys::{HRESULT, IID, NOERROR};
use log::{debug, error, LevelFilter};
use netidx::{
    path::Path,
    subscriber::{Event, Value},
};
use once_cell::sync::Lazy;
use simplelog;
use std::{
    boxed::Box, ffi::c_void, fs::File, mem, ptr, sync::mpsc, thread, time::Duration,
};
use windows::{
    core::{Abi, GUID},
    Win32::{
        Foundation::{SysAllocStringLen, PWSTR},
        System::{
            Com::{
                self, CoInitialize, CoUninitialize, IStream, ITypeInfo,
                Marshal::CoMarshalInterThreadInterfaceInStream,
                StructuredStorage::CoGetInterfaceAndReleaseStream, DISPPARAMS, EXCEPINFO,
                SAFEARRAY, SAFEARRAYBOUND, VARIANT, VARIANT_0_0_0,
            },
            Ole::{
                self, SafeArrayCreate, SafeArrayGetLBound, SafeArrayGetUBound,
                SafeArrayPutElement, VariantClear, VariantInit,
            },
            Threading::{CreateThread, THREAD_CREATION_FLAGS},
        },
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

// IRTDUpdateEvent is single apartment threaded, and that means we need to ask COM
// to make a proxy for us in order to run it in another thread. Since we MUST run in
// another thread to be async, this is mandatory. We have to marshal the interface when
// we receive it, and then unmarshal it in the update thread, which is then able to
// call into it.
struct IRTDUpdateEventThreadArgs {
    stream: IStream,
    rx: mpsc::Receiver<()>,
}

static IDISPATCH_GUID: GUID = GUID {
    data1: IID_IDISPATCH.data1,
    data2: IID_IDISPATCH.data2,
    data3: IID_IDISPATCH.data3,
    data4: IID_IDISPATCH.data4,
};

unsafe fn irtd_update_event_loop(
    update_notify: i32,
    rx: mpsc::Receiver<()>,
    idp: Com::IDispatch,
) {
    while let Ok(()) = rx.recv() {
        while let Ok(()) = rx.try_recv() {}
        loop {
            let mut args = [];
            let mut named_args = [];
            let mut params = DISPPARAMS {
                rgvarg: args.as_mut_ptr(),
                rgdispidNamedArgs: named_args.as_mut_ptr(),
                cArgs: 0,
                cNamedArgs: 0,
            };
            let mut result = Variant::null();
            let mut _arg_err = 0;
            let hr = idp.Invoke(
                update_notify,
                &GUID::zeroed(),
                0,
                Ole::DISPATCH_METHOD as u16,
                &mut params,
                result.as_mut_ptr(),
                ptr::null_mut(),
                &mut _arg_err,
            );
            match hr {
                Ok(()) => break,
                Err(e) => {
                    error!("IRTDUpdateEvent: update_notify failed {}", e);
                    thread::sleep(Duration::from_millis(250))
                }
            }
        }
    }
}

unsafe extern "system" fn irtd_update_event_thread(ptr: *mut c_void) -> u32 {
    maybe_init_logger();
    let args = Box::from_raw(ptr.cast::<IRTDUpdateEventThreadArgs>());
    match CoInitialize(ptr::null_mut()) {
        Ok(()) => (),
        Err(e) => {
            error!("update_event_thread: failed to initialize COM {}", e);
            return 0;
        }
    }
    let idp: Com::IDispatch = match CoGetInterfaceAndReleaseStream(&args.stream) {
        Ok(i) => i,
        Err(e) => {
            error!(
                "update_event_thread: failed to unmarshal the IDispatch interface {}",
                e
            );
            CoUninitialize();
            return 0;
        }
    };
    let mut update_notify = str_to_wstr("UpdateNotify");
    let mut dispid = 0x0;
    debug!("get_dispids: calling GetIDsOfNames");
    let hr = idp.GetIDsOfNames(
        &GUID::zeroed(),
        &PWSTR(update_notify.as_mut_ptr()),
        1,
        1000,
        &mut dispid,
    );
    debug!("update_event_thread: called GetIDsOfNames dispids: {:?}", dispid);
    if let Err(e) = hr {
        error!("update_event_thread: could not get names {}", e);
    }
    debug!("update_event_thread, init done, calling event loop");
    irtd_update_event_loop(dispid, args.rx, idp);
    CoUninitialize();
    0
}

pub(crate) struct IRTDUpdateEventWrap(mpsc::Sender<()>);

impl IRTDUpdateEventWrap {
    unsafe fn new(disp: Com::IDispatch) -> Result<Self> {
        let (tx, rx) = mpsc::channel();
        let stream = CoMarshalInterThreadInterfaceInStream(&IDISPATCH_GUID, disp)
            .map_err(|e| anyhow!(e.to_string()))?;
        let args = Box::new(IRTDUpdateEventThreadArgs { stream, rx });
        let mut threadid = 0u32;
        CreateThread(
            ptr::null_mut(),
            0,
            Some(irtd_update_event_thread),
            Box::into_raw(args).cast::<c_void>(),
            THREAD_CREATION_FLAGS::default(),
            &mut threadid,
        );
        Ok(IRTDUpdateEventWrap(tx))
    }

    pub(crate) fn update_notify(&self) {
        let _ = self.0.send(());
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

    unsafe fn get(&self, i: usize) -> &Variant {
        Variant::ref_from_raw((*self.0).rgvarg.offset(i as isize))
    }

    unsafe fn get_mut(&self, i: usize) -> &mut Variant {
        Variant::ref_from_raw_mut((*self.0).rgvarg.offset(i as isize))
    }
}

unsafe fn dispatch_server_start(server: &Server, params: *mut DISPPARAMS) -> Result<()> {
    let params = Params::new(params)?;
    server.server_start(IRTDUpdateEventWrap::new(params.get(0).try_into()?)?);
    Ok(())
}

unsafe fn dispatch_connect_data(server: &Server, params: *mut DISPPARAMS) -> Result<()> {
    let params = Params::new(params)?;
    if params.len() != 3 {
        bail!("wrong number of args")
    }
    let topic_id = TopicId(params.get(2).try_into()?);
    let topics: &SafeArray = params.get(1).try_into()?;
    let topics = topics.read()?;
    let path = match topics.iter()?.next() {
        None => bail!("not enough topics"),
        Some(v) => {
            let path: String = v.try_into()?;
            Path::from(path)
        }
    };
    Ok(server.connect_data(topic_id, path)?)
}

fn variant_of_event(e: Event) -> Variant {
    match e {
        Event::Unsubscribed => Variant::from("#SUB"),
        Event::Update(v) => match v {
            Value::I32(v) | Value::Z32(v) => Variant::from(v),
            Value::U32(v) | Value::V32(v) => Variant::from(v),
            Value::I64(v) | Value::Z64(v) => Variant::from(v),
            Value::U64(v) | Value::V64(v) => Variant::from(v),
            Value::F32(v) => Variant::from(v),
            Value::F64(v) => Variant::from(v),
            Value::True => Variant::from(true),
            Value::False => Variant::from(false),
            Value::String(s) => Variant::from(&*s),
            Value::Bytes(_) => Variant::from("#BIN"),
            Value::Null => Variant::null(),
            Value::Ok => Variant::from("OK"),
            Value::Error(e) => Variant::from(&format!("#ERR {}", &*e)),
            Value::Array(_) => Variant::from("#ARRAY"), // CR estokes: implement this
            Value::DateTime(d) => Variant::from(&d.to_string()),
            Value::Duration(d) => Variant::from(&format!("{}s", d.as_secs_f64())),
        },
    }
}

unsafe fn dispatch_refresh_data(
    server: &Server,
    params: *mut DISPPARAMS,
    result: &mut Variant,
) -> Result<()> {
    let params = Params::new(params)?;
    if params.len() != 1 {
        bail!("refresh_data unexpected number of params")
    }
    let ntopics = params.get_mut(0);
    let ntopics: &mut i32 = ntopics.try_into()?;
    let mut updates = server.refresh_data();
    let len = updates.len();
    *ntopics = len as i32;
    let mut array = SafeArray::new(&[
        SAFEARRAYBOUND { lLbound: 0, cElements: 2 },
        SAFEARRAYBOUND { lLbound: 0, cElements: len as u32 },
    ]);
    {
        let wh = array.write()?;
        for (i, (TopicId(tid), e)) in updates.drain().enumerate() {
            *wh.get_mut(&[0, i as i32])? = Variant::from(tid);
            *wh.get_mut(&[1, i as i32])? = variant_of_event(e);
        }
    }
    *result = Variant::from(array);
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
    let topic_id = TopicId(params.get(0).try_into()?);
    Ok(server.disconnect_data(topic_id))
}

com::class! {
    #[derive(Debug)]
    pub class NetidxRTD: IRTDServer(IDispatch) {
        server: Server,
    }

    impl IDispatch for NetidxRTD {
        fn get_type_info_count(&self, info: *mut u32) -> HRESULT {
            maybe_init_logger();
            debug!("get_type_info_count(info: {})", unsafe { *info });
            if !info.is_null() {
                unsafe { *info = 0; } // no we don't support type info
            }
            NOERROR
        }

        fn get_type_info(&self, _lcid: u32, _type_info: *mut *mut ITypeInfo) -> HRESULT { NOERROR }

        pub fn get_ids_of_names(
            &self,
            riid: *const IID,
            names: *const *mut u16,
            names_len: u32,
            lcid: u32,
            ids: *mut i32
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
            id: i32,
            iid: *const IID,
            lcid: u32,
            flags: u16,
            params: *mut DISPPARAMS,
            result: *mut VARIANT,
            exception: *mut EXCEPINFO,
            arg_error: *mut u32
        ) -> HRESULT {
            maybe_init_logger();
            debug!(
                "invoke(id: {}, iid: {:?}, lcid: {}, flags: {}, params: {:?}, result: {:?}, exception: {:?}, arg_error: {:?})",
                id, iid, lcid, flags, params, result, exception, arg_error
            );
            assert!(!params.is_null());
            let result = Variant::ref_from_raw_mut(result);
            match id {
                0 => {
                    debug!("ServerStart");
                    match dispatch_server_start(&self.server, params) {
                        Ok(()) => { *result = Variant::from(1); },
                        Err(e) => {
                            error!("server_start invalid arg {}", e);
                            *result = Variant::error();
                        }
                    }
               },
                1 => {
                    debug!("ServerTerminate");
                    self.server.server_terminate();
                    *result = Variant::from(1);
                },
                2 => {
                    debug!("ConnectData");
                    match dispatch_connect_data(&self.server, params) {
                        Ok(()) => { *result = Variant::from(1); },
                        Err(e) => {
                            error!("connect_data invalid arg {}", e);
                            *result = Variant::error();
                        }
                    }
                },
                3 => {
                    debug!("RefreshData");
                    match dispatch_refresh_data(&self.server, params, result) {
                        Ok(()) => (),
                        Err(e) => {
                            error!("refresh_data failed {}", e);
                            *result = Variant::error();
                        }
                    }
                },
                4 => {
                    debug!("DisconnectData");
                    match dispatch_disconnect_data(&self.server, params) {
                        Ok(()) => { *result = Variant::from(1); } 
                        Err(e) => {
                            error!("disconnect_data invalid arg {}", e);
                            *result = Variant::error()
                        }
                    }
                },
                5 => {
                    debug!("Heartbeat");
                    *result = Variant::from(1);
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
