#![allow(non_snake_case, non_upper_case_globals, non_camel_case_types)]
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        System::Com::*,
    },
};
use super::*;

// ClassFactory for Calculator
#[repr(C)]
#[implement(IClassFactory)]
pub struct CalculatorFactory;

impl CalculatorFactory {
    pub fn new() -> Self {
        Self
    }
}

impl IClassFactory_Impl for CalculatorFactory_Impl {
    fn CreateInstance(
        &self,
        outer: windows_core::Ref<'_, IUnknown>,
        iid: *const GUID,
        ppv: *mut *mut core::ffi::c_void,
    ) -> windows::core::Result<()> {
        if outer.is_some() {
            return Err(CLASS_E_NOAGGREGATION.into());
        }

        if iid.is_null() {
            return Err(E_POINTER.into());
        }

        if ppv.is_null() {
            return Err(E_POINTER.into());
        }

        // Create calculator with default settings
        // These will be overridden by Initialize call
        let calculator: Calculator = Calculator::new(2, "standard")?;
        let unknown: windows::core::IUnknown = calculator.into();

        unsafe {
            let res = unknown.query(iid, ppv);
            if res.is_err() {
                return Err(res.into());
            }
        }
        Ok(())
    }

    fn LockServer(&self, _lock: BOOL) -> windows::core::Result<()> {
        Ok(())
    }
}