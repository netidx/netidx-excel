use crate::comglue::{
    interface::IID_IDISPATCH,
    variant::{str_to_wstr, Variant},
    maybe_init_logger
};
use anyhow::{anyhow, Result};
use log::{debug, error};
use std::{boxed::Box, ffi::c_void, ptr, sync::mpsc, thread, time::Duration};
use windows::{
    core::GUID,
    Win32::{
        Foundation::PWSTR,
        System::{
            Com::{
                self, CoInitialize, CoUninitialize, IStream,
                Marshal::CoMarshalInterThreadInterfaceInStream,
                StructuredStorage::CoGetInterfaceAndReleaseStream, DISPPARAMS,
            },
            Ole,
            Threading::{CreateThread, THREAD_CREATION_FLAGS},
        },
    },
};

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

pub struct IRTDUpdateEventWrap(mpsc::Sender<()>);

impl IRTDUpdateEventWrap {
    pub unsafe fn new(disp: Com::IDispatch) -> Result<Self> {
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

    pub fn update_notify(&self) {
        let _ = self.0.send(());
    }
}
