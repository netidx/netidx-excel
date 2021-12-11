use crate::{
    comglue::{
        dispatch::IRTDUpdateEventWrap,
        interface::{IDispatch, IRTDServer, IRTDUpdateEvent},
        variant::{string_from_wstr, SafeArray, Variant},
    },
    server::{Server, TopicId},
};
use anyhow::{bail, Result};
use com::sys::{HRESULT, IID, NOERROR};
use log::{debug, error};
use netidx::{
    path::Path,
    subscriber::{Event, Value},
};
use windows::Win32::System::Com::{
    ITypeInfo, DISPPARAMS, EXCEPINFO, SAFEARRAY, SAFEARRAYBOUND, VARIANT,
};

struct Params(*mut DISPPARAMS);

impl Drop for Params {
    fn drop(&mut self) {
        unsafe {
            for i in 0..self.len() {
                if let Ok(v) = self.get_mut(i) {
                    *v = Variant::new();
                }
            }
        }
    }
}

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

    unsafe fn get(&self, i: usize) -> Result<&Variant> {
        if i < self.len() {
            Ok(Variant::ref_from_raw((*self.0).rgvarg.offset(i as isize)))
        } else {
            bail!("no param at index: {}", i)
        }
    }

    unsafe fn get_mut(&self, i: usize) -> Result<&mut Variant> {
        if i < self.len() {
            Ok(Variant::ref_from_raw_mut((*self.0).rgvarg.offset(i as isize)))
        } else {
            bail!("no param at index: {}", i)
        }
    }
}

unsafe fn dispatch_server_start(server: &Server, params: Params) -> Result<()> {
    server.server_start(IRTDUpdateEventWrap::new(params.get(0)?.try_into()?)?);
    Ok(())
}

unsafe fn dispatch_connect_data(server: &Server, params: Params) -> Result<()> {
    let topic_id = TopicId(params.get(2)?.try_into()?);
    let topics: &SafeArray = params.get(1)?.try_into()?;
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

fn variant_of_value(v: &Value) -> Variant {
    match v {
        Value::I32(v) | Value::Z32(v) => Variant::from(*v),
        Value::U32(v) | Value::V32(v) => Variant::from(*v),
        Value::I64(v) | Value::Z64(v) => Variant::from(*v),
        Value::U64(v) | Value::V64(v) => Variant::from(*v),
        Value::F32(v) => Variant::from(*v),
        Value::F64(v) => Variant::from(*v),
        Value::True => Variant::from(true),
        Value::False => Variant::from(false),
        Value::String(s) => Variant::from(&**s),
        Value::Bytes(_) => Variant::from("#BIN"),
        Value::Null => Variant::null(),
        Value::Ok => Variant::from("OK"),
        Value::Error(e) => Variant::from(&format!("#ERR {}", &**e)),
        Value::Array(_) => Variant::from(&format!("{}", v)),
        Value::DateTime(d) => Variant::from(&d.to_string()),
        Value::Duration(d) => Variant::from(&format!("{}s", d.as_secs_f64())),
    }
}

fn variant_of_event(e: &Event) -> Variant {
    match e {
        Event::Unsubscribed => Variant::from("#SUB"),
        Event::Update(v) => variant_of_value(v),
    }
}

unsafe fn dispatch_refresh_data(
    server: &Server,
    params: Params,
    result: &mut Variant,
) -> Result<()> {
    let ntopics = params.get_mut(0)?;
    let ntopics: &mut i32 = ntopics.try_into()?;
    let mut updates = server.refresh_data();
    let len = updates.len();
    *ntopics = len as i32;
    let mut array = SafeArray::new(&[
        SAFEARRAYBOUND { lLbound: 0, cElements: 2 },
        SAFEARRAYBOUND { lLbound: 0, cElements: len as u32 },
    ]);
    {
        let mut wh = array.write()?;
        for (i, (TopicId(tid), e)) in updates.drain().enumerate() {
            *wh.get_mut(&[0, i as i32])? = Variant::from(tid);
            *wh.get_mut(&[1, i as i32])? = variant_of_event(&e);
        }
    }
    *result = Variant::from(array);
    Ok(())
}

unsafe fn dispatch_disconnect_data(server: &Server, params: Params) -> Result<()> {
    let topic_id = TopicId(params.get(0)?.try_into()?);
    Ok(server.disconnect_data(topic_id))
}

com::class! {
    #[derive(Debug)]
    pub class NetidxRTD: IRTDServer(IDispatch) {
        server: Server,
    }

    impl IDispatch for NetidxRTD {
        fn get_type_info_count(&self, info: *mut u32) -> HRESULT {
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
            debug!(
                "invoke(id: {}, iid: {:?}, lcid: {}, flags: {}, params: {:?}, result: {:?}, exception: {:?}, arg_error: {:?})",
                id, iid, lcid, flags, params, result, exception, arg_error
            );
            assert!(!params.is_null());
            let result = Variant::ref_from_raw_mut(result);
            let params = match Params::new(params) {
                Ok(p) => p,
                Err(e) => {
                    error!("failed to wrap params {}", e);
                    *result = Variant::error();
                    return NOERROR;
                }
            };
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
            debug!("ServerStart called directly");
            NOERROR
        }

        fn connect_data(&self, _topic_id: i32, _topic: *const SAFEARRAY, _get_new_values: *mut VARIANT, _res: *mut VARIANT) -> HRESULT {
            debug!("ConnectData called directly");
            NOERROR
        }

        fn refresh_data(&self, _topic_count: *mut i32, _data: *mut SAFEARRAY) -> HRESULT {
            debug!("RefreshData called directly");
            NOERROR
        }

        fn disconnect_data(&self, _topic_id: i32) -> HRESULT {
            debug!("DisconnectData called directly");
            NOERROR
        }

        fn heartbeat(&self, _res: *mut i32) -> HRESULT {
            debug!("Heartbeat called directly");
            NOERROR
        }

        fn server_terminate(&self) -> HRESULT {
            debug!("ServerTerminate called directly");
            NOERROR
        }
    }
}
