use anyhow::{anyhow, bail, Error, Result};
use std::{
    convert::{From, TryInto},
    default::Default,
    ffi::{c_void, OsString},
    iter::Iterator,
    mem,
    ops::Drop,
    os::windows::ffi::{OsStrExt, OsStringExt},
    ptr,
};
use windows::{
    core::PCWSTR,
    Win32::{
        Foundation::SysAllocStringLen,
        Globalization::lstrlenW,
        System::{
            Com::{
                IDispatch, SAFEARRAY, SAFEARRAYBOUND, VARENUM, VARIANT, VARIANT_0_0_0,
                VT_ARRAY, VT_BOOL, VT_BSTR, VT_BYREF, VT_DISPATCH, VT_ERROR, VT_I4,
                VT_I8, VT_NULL, VT_R4, VT_R8, VT_UI4, VT_UI8, VT_VARIANT,
            },
            Ole::{
                SafeArrayCreate, SafeArrayDestroy, SafeArrayGetDim, SafeArrayGetLBound,
                SafeArrayGetUBound, SafeArrayGetVartype, SafeArrayLock,
                SafeArrayPtrOfIndex, SafeArrayUnlock, VariantClear, VariantInit,
            },
        },
    },
};

pub unsafe fn string_from_wstr<'a>(s: *mut u16) -> OsString {
    OsString::from_wide(std::slice::from_raw_parts(s, lstrlenW(PCWSTR(s)) as usize))
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
        Variant(unsafe { VariantInit() })
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
                if v.as_bool() {
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
            let s = unsafe { &*self.val().bstrVal };
            Ok(OsString::from_wide(s.as_wide()).to_string_lossy().to_string())
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
                match &*self.val().pdispVal {
                    None => bail!("null IDispatch interface"),
                    Some(d) => Ok(d.clone()),
                }
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
            Ok(unsafe {
                mem::transmute::<&*mut SAFEARRAY, &SafeArray>(&self.val().parray)
            })
        }
    }
}

impl<'a> TryInto<&'a mut SafeArray> for &'a mut Variant {
    type Error = Error;

    fn try_into(self) -> Result<&'a mut SafeArray, Self::Error> {
        if self.typ() != VARENUM(VT_ARRAY.0 | VT_VARIANT.0) {
            bail!("not a variant safearray")
        } else {
            Ok(unsafe {
                mem::transmute::<&mut *mut SAFEARRAY, &mut SafeArray>(
                    &mut self.val_mut().parray,
                )
            })
        }
    }
}

impl From<bool> for Variant {
    fn from(b: bool) -> Self {
        let mut v = Self::new();
        unsafe {
            v.set_typ(VT_BOOL);
            v.val_mut().boolVal.0 = if b { -1 } else { 0 };
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
            let bs = SysAllocStringLen(Some(&*str_to_wstr(s)));
            v.val_mut().bstrVal = mem::ManuallyDrop::new(bs);
            v
        }
    }
}

impl From<&String> for Variant {
    fn from(s: &String) -> Self {
        Variant::from(s.as_str())
    }
}

impl From<String> for Variant {
    fn from(s: String) -> Self {
        Variant::from(s.as_str())
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

    pub fn as_ptr(&self) -> *const VARIANT {
        unsafe { mem::transmute::<&Variant, &VARIANT>(self) as *const VARIANT }
    }

    pub fn as_mut_ptr(&mut self) -> *mut VARIANT {
        unsafe { mem::transmute::<&mut Variant, &mut VARIANT>(self) as *mut VARIANT }
    }

    // turn a const pointer to a `VARIANT` into a reference to a `Variant`.
    // take care to assign a reasonable lifetime.
    pub unsafe fn ref_from_raw<'a>(p: *const VARIANT) -> &'a Variant {
        mem::transmute::<*const VARIANT, &'a Variant>(p)
    }

    // turn a mut pointer to a `VARIANT` into a mutable reference to a `Variant`.
    // take care to assign a reasonable lifetime.
    pub unsafe fn ref_from_raw_mut<'a>(p: *mut VARIANT) -> &'a mut Variant {
        mem::transmute::<*mut VARIANT, &'a mut Variant>(p)
    }

    pub fn typ(&self) -> VARENUM {
        unsafe { self.0.Anonymous.Anonymous.vt }
    }

    unsafe fn set_typ(&mut self, typ: VARENUM) {
        (*self.0.Anonymous.Anonymous).vt = typ;
    }

    unsafe fn val(&self) -> &VARIANT_0_0_0 {
        &self.0.Anonymous.Anonymous.Anonymous
    }

    unsafe fn val_mut(&mut self) -> &mut VARIANT_0_0_0 {
        &mut (*self.0.Anonymous.Anonymous).Anonymous
    }
}

fn next_index(bounds: &[SAFEARRAYBOUND], idx: &mut [i32]) -> bool {
    let mut i = 0;
    while i < bounds.len() {
        if idx[i] < (bounds[i].lLbound + bounds[i].cElements as i32) {
            idx[i] += 1;
            for j in 0..i {
                idx[j] = bounds[j].lLbound;
            }
            break;
        }
        i += 1;
    }
    i < bounds.len()
}

pub struct SafeArrayIterMut<'a> {
    array: &'a mut SafeArray,
    bounds: Vec<SAFEARRAYBOUND>,
    idx: Vec<i32>,
    end: bool,
}

impl<'a> Iterator for SafeArrayIterMut<'a> {
    type Item = &'a mut Variant;

    fn next(&mut self) -> Option<Self::Item> {
        if self.end {
            None
        } else {
            let res = unsafe {
                let mut vp: *mut VARIANT = ptr::null_mut();
                SafeArrayPtrOfIndex(
                    self.array.0,
                    self.idx.as_ptr(),
                    &mut vp as *mut *mut VARIANT as *mut *mut c_void,
                )
                .ok()?;
                Some(Variant::ref_from_raw_mut(vp))
            };
            self.end = next_index(&self.bounds, &mut self.idx);
            res
        }
    }
}

pub struct SafeArrayIter<'a> {
    array: &'a SafeArray,
    bounds: Vec<SAFEARRAYBOUND>,
    idx: Vec<i32>,
    end: bool,
}

impl<'a> Iterator for SafeArrayIter<'a> {
    type Item = &'a Variant;

    fn next(&mut self) -> Option<Self::Item> {
        if self.end {
            None
        } else {
            let res = unsafe {
                let mut vp: *mut VARIANT = ptr::null_mut();
                SafeArrayPtrOfIndex(
                    self.array.0,
                    self.idx.as_ptr(),
                    &mut vp as *mut *mut VARIANT as *mut *mut c_void,
                )
                .ok()?;
                Some(Variant::ref_from_raw(vp))
            };
            self.end = next_index(&self.bounds, &mut self.idx);
            res
        }
    }
}

pub struct SafeArrayReadGuard<'a>(&'a SafeArray);

impl<'a> Drop for SafeArrayReadGuard<'a> {
    fn drop(&mut self) {
        unsafe {
            let _ = SafeArrayUnlock(self.0 .0);
        }
    }
}

impl<'a> SafeArrayReadGuard<'a> {
    pub fn dims(&self) -> u32 {
        self.0.dims()
    }

    pub fn bound(&self, dim: u32) -> Result<SAFEARRAYBOUND> {
        self.0.bound(dim)
    }

    pub fn bounds(&self) -> Result<Vec<SAFEARRAYBOUND>> {
        self.0.bounds()
    }

    pub fn iter(&self) -> Result<SafeArrayIter> {
        let bounds = self.bounds()?;
        let idx =
            (0..bounds.len()).into_iter().map(|i| bounds[i].lLbound).collect::<Vec<_>>();
        Ok(SafeArrayIter { array: self.0, bounds, idx, end: false })
    }

    pub fn get(&self, idx: &[i32]) -> Result<&Variant> {
        self.0.get(idx)
    }
}

pub struct SafeArrayWriteGuard<'a>(&'a mut SafeArray);

impl<'a> Drop for SafeArrayWriteGuard<'a> {
    fn drop(&mut self) {
        unsafe {
            let _ = SafeArrayUnlock(self.0 .0);
        }
    }
}

impl<'a> SafeArrayWriteGuard<'a> {
    pub fn dims(&self) -> u32 {
        self.0.dims()
    }

    pub fn bound(&self, dim: u32) -> Result<SAFEARRAYBOUND> {
        self.0.bound(dim)
    }

    pub fn bounds(&self) -> Result<Vec<SAFEARRAYBOUND>> {
        self.0.bounds()
    }

    pub fn iter(&self) -> Result<SafeArrayIter> {
        let bounds = self.bounds()?;
        let idx =
            (0..bounds.len()).into_iter().map(|i| bounds[i].lLbound).collect::<Vec<_>>();
        Ok(SafeArrayIter { array: self.0, bounds, idx, end: false })
    }

    pub fn iter_mut(&mut self) -> Result<SafeArrayIterMut> {
        let bounds = self.bounds()?;
        let idx =
            (0..bounds.len()).into_iter().map(|i| bounds[i].lLbound).collect::<Vec<_>>();
        Ok(SafeArrayIterMut { array: self.0, bounds, idx, end: false })
    }

    pub fn get(&self, idx: &[i32]) -> Result<&Variant> {
        self.0.get(idx)
    }

    pub fn get_mut(&mut self, idx: &[i32]) -> Result<&mut Variant> {
        self.0.get_mut(idx)
    }
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
        let t =
            unsafe { SafeArrayCreate(VT_VARIANT, bounds.len() as u32, bounds.as_ptr()) };
        SafeArray(t)
    }

    unsafe fn check_pointer(p: *const SAFEARRAY) -> Result<()> {
        let typ = SafeArrayGetVartype(p)
            .map_err(|e| anyhow!("couldn't get safearray type {}", e.to_string()))?;
        if typ != VT_VARIANT {
            bail!("not a variant array")
        }
        Ok(())
    }

    pub unsafe fn from_raw<'a>(p: *mut SAFEARRAY) -> Result<Self> {
        Self::check_pointer(p)?;
        Ok(mem::transmute::<*mut SAFEARRAY, SafeArray>(p))
    }

    pub fn write<'a>(&'a mut self) -> Result<SafeArrayWriteGuard<'a>> {
        unsafe {
            SafeArrayLock(self.0)
                .map_err(|e| anyhow!("failed to lock safearray {}", e.to_string()))?;
            Ok(SafeArrayWriteGuard(self))
        }
    }

    pub fn read<'a>(&'a self) -> Result<SafeArrayReadGuard<'a>> {
        unsafe {
            SafeArrayLock(self.0)
                .map_err(|e| anyhow!("failed to lock safearray {}", e.to_string()))?;
            Ok(SafeArrayReadGuard(self))
        }
    }

    fn dims(&self) -> u32 {
        unsafe { SafeArrayGetDim(self.0) }
    }

    fn bound(&self, dim: u32) -> Result<SAFEARRAYBOUND> {
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

    fn bounds(&self) -> Result<Vec<SAFEARRAYBOUND>> {
        let dims = self.dims();
        let mut res = Vec::with_capacity(dims as usize);
        for i in 1..=dims {
            let bound = self.bound(i)?;
            res.push(bound)
        }
        Ok(res)
    }

    fn get(&self, idx: &[i32]) -> Result<&Variant> {
        unsafe {
            let mut vp: *mut VARIANT = ptr::null_mut();
            match SafeArrayPtrOfIndex(
                self.0,
                idx.as_ptr(),
                &mut vp as *mut *mut VARIANT as *mut *mut c_void,
            ) {
                Ok(()) => Ok(Variant::ref_from_raw(vp)),
                Err(e) => bail!("could not access idx: {:?}, {}", idx, e.to_string()),
            }
        }
    }

    fn get_mut(&mut self, idx: &[i32]) -> Result<&mut Variant> {
        unsafe {
            let mut vp: *mut VARIANT = ptr::null_mut();
            match SafeArrayPtrOfIndex(
                self.0,
                idx.as_ptr(),
                &mut vp as *mut *mut VARIANT as *mut *mut c_void,
            ) {
                Ok(()) => Ok(Variant::ref_from_raw_mut(vp)),
                Err(e) => bail!("could not access idx: {:?}, {}", idx, e.to_string()),
            }
        }
    }
}
