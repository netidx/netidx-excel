use anyhow::{anyhow, bail, Error, Result};
use std::{
    convert::{From, TryInto},
    default::Default,
    ffi::{c_void, OsString},
    mem,
    ops::{Deref, DerefMut, Drop},
    os::windows::ffi::{OsStrExt, OsStringExt},
    ptr,
};
use windows::{
    core::Abi,
    Win32::{
        Foundation::{SysAllocStringLen, PWSTR},
        Globalization::lstrlenW,
        System::{
            Com::{IDispatch, SAFEARRAY, SAFEARRAYBOUND, VARIANT, VARIANT_0_0_0},
            Ole::{
                SafeArrayCreate, SafeArrayDestroy, SafeArrayGetDim,
                SafeArrayGetElement, SafeArrayGetLBound, SafeArrayGetUBound,
                SafeArrayGetVartype, SafeArrayPutElement, VariantClear, VariantInit,
                VARENUM, VT_BOOL, VT_BSTR, VT_BYREF, VT_DISPATCH, VT_ERROR, VT_I4, VT_I8,
                VT_NULL, VT_R4, VT_R8, VT_UI4, VT_UI8, VT_VARIANT, VT_ARRAY
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
                Ok(IDispatch::from_abi(self.val().pdispVal)
                    .map_err(|e| anyhow!(e.to_string()))?)
            }
        }
    }
}

impl<'a> TryInto<&'a SafeArray> for &'a Variant {
    type Error = Error;

    fn try_into(self) -> Result<&'a SafeArray, Self::Error> {
        if self.typ() != VARENUM(VT_ARRAY.0 | VT_VARIANT.0) {
            bail!("not a variant safearray")
        } else {
            Ok(unsafe { SafeArray::from_raw(self.val().parray)? })
        }
    }
}

impl<'a> TryInto<&'a mut SafeArray> for &'a mut Variant {
    type Error = Error;

    fn try_into(self) -> Result<&'a mut SafeArray, Self::Error> {
        if self.typ() != VARENUM(VT_ARRAY.0 | VT_VARIANT.0) {
            bail!("not a variant safearray")
        } else {
            Ok(unsafe { SafeArray::from_raw_mut(self.val().parray)? })
        }
    }
}

impl From<bool> for Variant {
    fn from(b: bool) -> Self {
        let mut v = Self::new();
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

impl From<u64> for Variant {
    fn from(i: u64) -> Self {
        let mut v = Self::new();
        unsafe {
            v.set_typ(VT_UI8);
            v.val_mut().ullVal = i;
            v
        }
    }
}

impl From<f32> for Variant {
    fn from(i: f32) -> Self {
        let mut v = Self::new();
        unsafe {
            v.set_typ(VT_R4);
            v.val_mut().fltVal = i;
            v
        }
    }
}

impl From<f64> for Variant {
    fn from(i: f64) -> Self {
        let mut v = Self::new();
        unsafe {
            v.set_typ(VT_R8);
            v.val_mut().dblVal = i;
            v
        }
    }
}

impl From<&str> for Variant {
    fn from(s: &str) -> Self {
        let mut v = Self::new();
        unsafe {
            v.set_typ(VT_BSTR);
            let mut s = str_to_wstr(s);
            let bs = SysAllocStringLen(PWSTR(s.as_mut_ptr()), s.len() as u32);
            v.val_mut().bstrVal = mem::ManuallyDrop::new(bs);
            v
        }
    }
}

impl From<SafeArray> for Variant {
    fn from(a: SafeArray) -> Self {
        let mut v = Self::new();
        unsafe {
            v.set_typ(VARENUM(VT_ARRAY.0 | VT_VARIANT.0));
            v.val_mut().parray = a.0;
            // the variant is now responsible for deallocating the safe array
            mem::forget(a); 
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
    pub unsafe fn from_raw<'a>(p: *const VARIANT) -> &'a Variant {
        mem::transmute::<&'a VARIANT, &'a Variant>(&*p)
    }

    // turn a mut pointer to a `VARIANT` into a mutable reference to a `Variant`.
    // take care to assign a reasonable lifetime.
    pub unsafe fn from_raw_mut<'a>(p: *mut VARIANT) -> &'a mut Variant {
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

    /*
    fn set_safearray(&mut self, v: *mut SAFEARRAY) {
        VariantInit(self.0);
        self.set_typ(Ole::VARENUM(Ole::VT_ARRAY.0 | Ole::VT_VARIANT.0));
        self.val_mut().parray = v;
    }
    */
}

#[repr(transparent)]
pub struct SafeArray(*mut SAFEARRAY);

impl Drop for SafeArray {
    fn drop(&mut self) {
        unsafe {
            let _ = SafeArrayDestroy(self.0);
        }
    }
}

impl SafeArray {
    pub fn new(bounds: &[SAFEARRAYBOUND]) -> SafeArray {
        let t = unsafe {
            SafeArrayCreate(VT_VARIANT.0 as u16, bounds.len() as u32, bounds.as_ptr())
        };
        SafeArray(t)
    }

    unsafe fn check_ptr(p: *const SAFEARRAY) -> Result<()> {
        let typ = SafeArrayGetVartype(p)
            .map_err(|e| anyhow!("couldn't get safearray type {}", e.to_string()))?;
        if typ != VT_VARIANT.0 as u16 {
            bail!("not a variant array")
        }
        Ok(())
    }

    pub unsafe fn from_raw<'a>(p: *const SAFEARRAY) -> Result<&'a Self> {
        Self::check_ptr(p)?;
        Ok(mem::transmute::<&SAFEARRAY, &SafeArray>(&*p))
    }

    pub unsafe fn from_raw_mut<'a>(p: *mut SAFEARRAY) -> Result<&'a mut Self> {
        Self::check_ptr(p)?;
        Ok(mem::transmute::<&mut SAFEARRAY, &mut SafeArray>(&mut *p))
    }

    pub fn dims(&self) -> u32 {
        unsafe { SafeArrayGetDim(self.0) }
    }

    pub fn bound(&self, dim: u32) -> Result<SAFEARRAYBOUND> {
        unsafe {
            let lbound = SafeArrayGetLBound(self.0, dim).map_err(|e| {
                anyhow!("couldn't get safe array lower bound {}", e.to_string())
            })?;
            let ubound = SafeArrayGetUBound(self.0, dim).map_err(|e| {
                anyhow!("couldn't get safe array upper bound {}", e.to_string())
            })?;
            Ok(SAFEARRAYBOUND {
                cElements: (1 + ubound - lbound) as u32,
                lLbound: lbound,
            })
        }
    }

    fn get(&self, idx: &[i32]) -> Result<Variant> {
        unsafe {
            let mut res = Variant::new();
            SafeArrayGetElement(
                self.0,
                idx.as_ptr(),
                &mut res as *mut Variant as *mut c_void,
            )
            .map_err(|e| anyhow!("failed to get safe array element {}", e.to_string()))?;
            Ok(res)
        }
    }

    fn set(&mut self, idx: &[i32], v: &Variant) -> Result<()> {
        unsafe {
            SafeArrayPutElement(
                self.0,
                idx.as_ptr(),
                v as *const Variant as *const c_void,
            )
            .map_err(|e| anyhow!("failed to set safearray element {}", e.to_string()))?;
            Ok(())
        }
    }
}
