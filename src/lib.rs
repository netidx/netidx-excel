#[macro_use]
extern crate serde_derive;
mod comglue;
mod server;
use anyhow::{bail, Result};
use com::{
    production::Class,
    sys::{CLASS_E_CLASSNOTAVAILABLE, CLSID, HRESULT, IID, NOERROR, SELFREG_E_CLASS},
};
use comglue::glue::NetidxRTD;
use comglue::interface::CLSID;
use std::{ffi::c_void, mem, ptr};

// sadly this doesn't register the class name, just the ID, so we must do all the
// registration ourselves because excel requires the name to be mapped to the id
//com::inproc_dll_module![(CLSID, NetidxRTD),];

static mut _HMODULE: *mut c_void = ptr::null_mut();

#[no_mangle]
unsafe extern "system" fn DllMain(
    hinstance: *mut c_void,
    fdw_reason: u32,
    _reserved: *mut c_void,
) -> i32 {
    const DLL_PROCESS_ATTACH: u32 = 1;
    if fdw_reason == DLL_PROCESS_ATTACH {
        _HMODULE = hinstance;
    }
    1
}

#[no_mangle]
unsafe extern "system" fn DllGetClassObject(
    class_id: *const CLSID,
    iid: *const IID,
    result: *mut *mut c_void,
) -> HRESULT {
    assert!(
        !class_id.is_null(),
        "class id passed to DllGetClassObject should never be null"
    );

    let class_id = &*class_id;
    if class_id == &CLSID {
        let instance = <NetidxRTD as Class>::Factory::allocate();
        instance.QueryInterface(&*iid, result)
    } else {
        CLASS_E_CLASSNOTAVAILABLE
    }
}

use winreg::{enums::*, RegKey};

extern "system" {
    fn GetModuleFileNameA(hModule: *mut c_void, lpFilename: *mut i8, nSize: u32) -> u32;
}

unsafe fn get_dll_file_path(hmodule: *mut c_void) -> String {
    const MAX_FILE_PATH_LENGTH: usize = 260;

    let mut path = [0u8; MAX_FILE_PATH_LENGTH];

    let len = GetModuleFileNameA(
        hmodule,
        path.as_mut_ptr() as *mut _,
        MAX_FILE_PATH_LENGTH as _,
    );

    String::from_utf8(path[..len as usize].to_vec()).unwrap()
}

fn clsid(id: CLSID) -> String {
    format!("{{{}}}", id)
}

fn register_clsid(root: &RegKey, clsid: &String) -> Result<()> {
    let (by_id, _) = root.create_subkey(&format!("CLSID\\{}", &clsid))?;
    let (by_id_inproc, _) = by_id.create_subkey("InprocServer32")?;
    by_id.set_value(&"", &"NetidxRTD")?;
    by_id_inproc.set_value("", &unsafe { get_dll_file_path(_HMODULE) })?;
    Ok(())
}

fn dll_register_server() -> Result<()> {
    let hkcr = RegKey::predef(HKEY_CLASSES_ROOT);
    let (by_name, _) = hkcr.create_subkey("NetidxRTD\\CLSID")?;
    let clsid = clsid(CLSID);
    by_name.set_value("", &clsid)?;
    if mem::size_of::<usize>() == 8 {
        register_clsid(&hkcr, &clsid)?;
    } else if mem::size_of::<usize>() == 4 {
        let wow = hkcr.open_subkey("WOW6432Node")?;
        register_clsid(&wow, &clsid)?;
    } else {
        bail!("can't figure out the word size")
    }
    Ok(())
}

#[no_mangle]
extern "system" fn DllRegisterServer() -> HRESULT {
    match dll_register_server() {
        Err(_) => SELFREG_E_CLASS,
        Ok(()) => NOERROR,
    }
}

fn dll_unregister_server() -> Result<()> {
    let hkcr = RegKey::predef(HKEY_CLASSES_ROOT);
    let clsid = clsid(CLSID);
    hkcr.delete_subkey_all("NetidxRTD")?;
    assert!(clsid.len() > 0);
    if mem::size_of::<usize>() == 8 {
        hkcr.delete_subkey_all(&format!("CLSID\\{}", clsid))?;
    } else if mem::size_of::<usize>() == 4 {
        hkcr.delete_subkey_all(&format!("WOW6432Node\\CLSID\\{}", clsid))?;
    } else {
        bail!("could not determine the word size")
    }
    Ok(())
}

#[no_mangle]
extern "system" fn DllUnregisterServer() -> HRESULT {
    match dll_unregister_server() {
        Err(_) => SELFREG_E_CLASS,
        Ok(()) => NOERROR,
    }
}
