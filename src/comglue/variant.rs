use std::{
    default::Default,
    ffi::OsString,
    mem,
    ops::{Deref, DerefMut, Drop},
    os::windows::ffi::{OsStrExt, OsStringExt},
    convert::{TryInto, From},
};
use anyhow::{bail, anyhow, Result, Error};
use windows::{
    core::Abi,
    Win32::{
        Globalization::lstrlenW,
        Foundation::{SysAllocStringLen, PWSTR},
        System::{
            Com::{IDispatch, SAFEARRAY, SAFEARRAYBOUND, VARIANT, VARIANT_0_0_0},
            Ole::{
                SafeArrayCreate, SafeArrayCreateVector, SafeArrayGetElement, SafeArrayGetLBound,
                SafeArrayGetUBound, SafeArrayPutElement, VariantClear, VariantInit,
                VARENUM, VT_BOOL, VT_BSTR, VT_ERROR, VT_I4, VT_I8, VT_NULL, VT_R4, VT_R8,
                VT_UI4, VT_UI8, VT_BYREF, VT_DISPATCH, VarBoolFromCy,
            },
        },
    },
};

pub unsafe fn string_from_wstr<'a>(s: *mut u16) -> OsString {
    OsString::from_wide(std::slice::from_raw_parts(s, lstrlenW(PWSTR(s)) as usize))
}

pub fn str_to_wstr(s: &str) -> Vec<u16> {
    let mut v = OsString::from(s).encode_wide().collect::<Vec<_>>();
    v.push(0);
    v
}

#[repr(transparent)]
pub struct Variant(VARIANT);

impl Default for Variant {
    fn default() -> Self {
        Variant(unsafe {
            let mut v = mem::zeroed();
            VariantInit(&mut v);
            v
        })
    }
}

impl Drop for Variant {
    fn drop(&mut self) {
        let _ = unsafe { VariantClear(&mut self.0) };
    }
}

impl<'a> TryInto<bool> for &'a Variant {
    type Error = Error;

    fn try_into(self) -> Result<bool, Self::Error> {
        if self.typ() != VT_BOOL {
            bail!("not a bool")
        } else {
            unsafe {
                let v = self.val().boolVal;
                if v == -1 {
                    Ok(true)
                } else {
                    Ok(false)
                }
            }
        }
    }
}

impl<'a> TryInto<i32> for &'a Variant {
    type Error = Error;

    fn try_into(self) -> Result<i32, Self::Error> {
        if self.typ() != VT_I4 {
            bail!("not an i32")
        } else {
            unsafe { Ok(self.val().lVal) }
        }
    }
}

impl<'a> TryInto<&'a mut i32> for &'a mut Variant {
    type Error = Error;

    fn try_into(self) -> Result<&'a mut i32, Self::Error> {
        if self.typ() != VARENUM(VT_I4.0 | VT_BYREF.0) {
            bail!("not a byref i32")
        } else {
            Ok(unsafe { &mut *self.val().plVal })
        }
    }
}

impl<'a> TryInto<String> for &'a Variant {
    type Error = Error;

    fn try_into(self) -> Result<String, Self::Error> {
        if self.typ() != VT_BSTR {
            bail!("not a string")
        } else {
            unsafe {
                let s = *self.val().bstrVal;
                Ok(string_from_wstr(s.0).to_string_lossy().to_string())
            }
        }
    }
}

impl<'a> TryInto<IDispatch> for &'a Variant {
    type Error = Error;

    fn try_into(self) -> Result<IDispatch, Self::Error> {
        if self.typ() != VT_DISPATCH {
            bail!("not an IDispatch interface")
        } else {
            unsafe {
                Ok(IDispatch::from_abi(self.val().pdispVal).map_err(|e| anyhow!(e.to_string()))?)
            }
        }
    }
}

impl From<bool> for Variant {
    fn from(b: bool) -> Self {
        let mut v = Self::net();
        unsafe {
            v.set_typ(VT_BOOL);
            v.val_mut().boolVal = if b { -1 } else { 0 };
            v
        }
    }
}

impl From<i32> for Variant {
    fn from(i: i32) -> Self {
        let mut v = Self::new();
        unsafe {
            v.set_typ(VT_I4);
            v.val_mut().lVal = i;
            v
        }
    }
}

impl From<u32> for Variant {
    fn from(i: u32) -> Self {
        let mut v = Self::new();
        unsafe {
            v.set_typ(VT_UI4);
            v.val_mut().ulVal = i;
            v
        }
    }
}

impl From<i64> for Variant {
    fn from(i: i64) -> Self {
        let mut v = Self::new();
        unsafe {
            v.set_typ(VT_I8);
            v.val_mut().llVal = i;
            v
        }
    }
}

impl Variant {
    pub fn new() -> Variant {
        Self::default()
    }

    pub fn null() -> Variant {
        let mut v = Self::default();
        unsafe { v.set_typ(VT_NULL) }
        v
    }

    pub fn error() -> Variant {
        let mut v = Self::default();
        unsafe { v.set_typ(VT_ERROR) }
        v
    }

    // turn a const pointer to a `VARIANT` into a reference to a `Variant`.
    // take care to assign a reasonable lifetime.
    pub unsafe fn from_ref_raw<'a>(p: *const VARIANT) -> &'a Variant {
        mem::transmute::<&'a VARIANT, &'a Variant>(&*p)
    }

    // turn a mut pointer to a `VARIANT` into a mutable reference to a `Variant`.
    // take care to assign a reasonable lifetime.
    pub unsafe fn from_ref_raw_mut<'a>(p: *mut VARIANT) -> &'a mut Variant {
        mem::transmute::<&'a mut VARIANT, &'a mut Variant>(&mut *p)
    }

    pub fn typ(&self) -> VARENUM {
        VARENUM(unsafe { self.0.Anonymous.Anonymous.vt as i32 })
    }

    unsafe fn set_typ(&mut self, typ: VARENUM) {
        (*self.0.Anonymous.Anonymous).vt = typ.0 as u16;
    }

    unsafe fn val(&self) -> &VARIANT_0_0_0 {
        &self.0.Anonymous.Anonymous.Anonymous
    }

    unsafe fn val_mut(&mut self) -> &mut VARIANT_0_0_0 {
        &mut (*self.0.Anonymous.Anonymous).Anonymous
    }

    fn set_bool(&mut self, v: bool) {
        unsafe {
            self.set_typ(VT_BOOL);
            self.val_mut().boolVal = if v { -1 } else { 0 };
        }
    }

    fn set_null(&mut self) {
        unsafe {
            self.set_typ(VT_NULL);
            self.val_mut().llVal = 0;
        }
    }

    fn set_error(&mut self) {
        unsafe {
            self.set_typ(VT_ERROR);
            self.val_mut().llVal = 0;
        }
    }

    fn set_i32(&mut self, v: i32) {
        unsafe {
            self.set_typ(VT_I4);
            self.val_mut().lVal = v;
        }
    }

    fn set_u32(&mut self, v: u32) {
        unsafe {
            self.set_typ(VT_UI4);
            self.val_mut().ulVal = v;
        }
    }

    fn set_i64(&mut self, v: i64) {
        unsafe {
            self.clear();
            self.set_typ(VT_I8);
            self.val_mut().llVal = v;
        }
    }

    fn set_u64(&mut self, v: u64) {
        unsafe {
            self.clear();
            self.set_typ(VT_UI8);
            self.val_mut().ullVal = v;
        }
    }

    fn set_f32(&mut self, v: f32) {
        unsafe {
            self.clear();
            self.set_typ(VT_R4);
            self.val_mut().fltVal = v;
        }
    }

    fn set_f64(&mut self, v: f64) {
        unsafe {
            self.clear();
            self.set_typ(VT_R8);
            self.val_mut().dblVal = v;
        }
    }

    fn set_string(&mut self, v: &str) {
        unsafe {
            self.clear();
            self.set_typ(VT_BSTR);
            let mut s = str_to_wstr(v);
            let bs = SysAllocStringLen(PWSTR(s.as_mut_ptr()), s.len() as u32);
            self.val_mut().bstrVal = mem::ManuallyDrop::new(bs);
            self.set_typ(VT_BSTR);
            let mut s = str_to_wstr(v);
            let bs = SysAllocStringLen(PWSTR(s.as_mut_ptr()), s.len() as u32);
            self.val_mut().bstrVal = mem::ManuallyDrop::new(bs);
        }
    }

    /*
    fn set_safearray(&mut self, v: *mut SAFEARRAY) {
        VariantInit(self.0);
        self.set_typ(Ole::VARENUM(Ole::VT_ARRAY.0 | Ole::VT_VARIANT.0));
        self.val_mut().parray = v;
    }
    */

    fn get_i32(&self) -> Result<i32> {
        if self.typ() == VT_I4 {
            Ok(unsafe { self.val().lVal })
        } else {
            bail!("not a i32 value")
        }
    }

    unsafe fn get_byref_i32<'a>(&'a self) -> Result<&'a mut i32> {
        if self.typ() == VARENUM(VT_I4.0 | VT_BYREF.0) {
            Ok(unsafe { &mut *self.val().plVal })
        } else {
            bail!("not a byref i32 value")
        }
    }

    fn get_string(&self) -> Result<String> {
        if self.typ() == VT_BSTR {
            unsafe {
                let s = *self.val().bstrVal;
                Ok(string_from_wstr(s.0).to_string_lossy().to_string())
            }
        } else {
            bail!("not a string value")
        }
    }

    unsafe fn get_variant_vector(&self) -> Result<VariantVector> {
        if self.typ() == (Ole::VT_ARRAY.0 | Ole::VT_VARIANT.0) as u16 {
            Ok(VariantVector::new(self.val().parray))
        } else {
            bail!("not a variant array")
        }
    }
}

#[repr(transparent)]
pub struct VariantArray(*mut SAFEARRAY);

impl VariantArray {
    fn new(len: usize) -> VariantArray {

    }
}