#![allow(non_snake_case, non_upper_case_globals, non_camel_case_types)]
use std::ffi::CString;

use windows::{
    core::*,
    Win32::{
        Foundation::*,
        System::Registry::*,
    },
};
mod register_utils;
mod class_factory;
use super::*;

#[no_mangle]
pub extern "system" fn DllGetClassObject(
    clsid: *const GUID,
    iid: *const GUID,
    ppv: *mut *mut core::ffi::c_void,
) -> HRESULT {
    if clsid.is_null() {
        return E_POINTER;
    }

    if iid.is_null() {
        return E_POINTER;
    }

    if ppv.is_null() {
        return E_POINTER;
    }

    let factory: class_factory::CalculatorFactory = class_factory::CalculatorFactory::new();
    let unknown: windows::core::IUnknown = factory.into();
    unsafe {
        if *clsid != CLSID_Calculator {
            return CLASS_E_CLASSNOTAVAILABLE;
        }
        unknown.query(iid, ppv)
    }
}

#[no_mangle]
pub extern "system" fn DllRegisterServer() -> HRESULT {
    fn register_com_object() -> windows::core::Result<()> {
        unsafe {
            // 1. Register ProgID
            let mut prog_key = HKEY::default();
            let c_subkey = CString::new(ProgID).unwrap();
            let win32_error = RegCreateKeyA(
                HKEY_CLASSES_ROOT,
                PCSTR::from_raw(c_subkey.as_ptr() as *const u8),
                &mut prog_key,
            );
            if win32_error.is_err() {
                return Err(win32_error.into());
            }

            // Set default value for ProgID
            let win32_error = RegSetValueExA(
                prog_key,
                None,
                None,
                REG_SZ,
                Some("Rust Calculator COM Object".as_bytes()),
            );

            if win32_error.is_err() {
                RegCloseKey(prog_key).0;
                return Err(win32_error.into());
            }

            // Add CLSID key under ProgID
            let mut clsid_prog_key = HKEY::default();
            let win32_error = RegCreateKeyA(
                prog_key,
                s!("CLSID"),
                &mut clsid_prog_key
            );
            if win32_error.is_err() {
                RegCloseKey(prog_key).0;
                return Err(win32_error.into());
            }

            let clsid_string = format!("{{{:?}}}", CLSID_Calculator);
            let win32_error = RegSetValueExA(
                clsid_prog_key,
                None,
                None,
                REG_SZ,
                Some(clsid_string.as_bytes()),
            );
            if win32_error.is_err() {
                RegCloseKey(clsid_prog_key).0;
                RegCloseKey(prog_key).0;
                return Err(win32_error.into());
            }

            RegCloseKey(clsid_prog_key).0;
            RegCloseKey(prog_key).0;

            // 2. Register CLSID
            let clsid_key_path = format!("CLSID\\{{{:?}}}", CLSID_Calculator);
            let c_path = CString::new(clsid_key_path).unwrap();
            let mut clsid_key = HKEY::default();
            let win32_error = RegCreateKeyA(
                HKEY_CLASSES_ROOT,
                PCSTR(c_path.as_ptr() as *const u8),
                &mut clsid_key,
            );
            if win32_error.is_err() {
                return Err(win32_error.into());
            }

            // Set default value
            let win32_error = RegSetValueExA(
                clsid_key,
                None,
                None,
                REG_SZ,
                Some("Rust Calculator COM Object".as_bytes()),
            );
            if win32_error.is_err() {
                RegCloseKey(clsid_key).0;
                return Err(win32_error.into());
            }

            // Add ProgID under CLSID
            let mut prog_id_key = HKEY::default();
            let win32_error = RegCreateKeyA(
                clsid_key,
                s!("ProgID"),
                &mut prog_id_key
            );
            if win32_error.is_err() {
                RegCloseKey(clsid_key).0;
                return Err(win32_error.into());
            }

            let win32_error = RegSetValueExA(
                prog_id_key,
                None,
                None,
                REG_SZ,
                Some(ProgID.as_bytes()),
            );
            if win32_error.is_err() {
                RegCloseKey(prog_id_key).0;
                RegCloseKey(clsid_key).0;
                return Err(win32_error.into());
            }

            // Set InprocServer32
            let mut inproc_key = HKEY::default();
            let win32_error = RegCreateKeyA(
                clsid_key,
                s!("InprocServer32"),
                &mut inproc_key
            );
            if win32_error.is_err() {
                RegCloseKey(prog_id_key).0;
                RegCloseKey(clsid_key).0;
                return Err(win32_error.into());
            }

            let path = register_utils::get_this_module_path()?;
            let path_u8: Vec<u8> = path.iter().map(|&x| x as u8).collect::<Vec<u8>>();
            let win32_error = RegSetValueExA(
                inproc_key,
                None,
                None,
                REG_SZ,
                Some(path_u8.as_slice()),
            );
            if win32_error.is_err() {
                RegCloseKey(inproc_key).0;
                RegCloseKey(prog_id_key).0;
                RegCloseKey(clsid_key).0;
                return Err(win32_error.into());
            }

            let win32_error = RegSetValueExA(
                inproc_key,
                s!("ThreadingModel"),
                None,
                REG_SZ,
                Some("Both".as_bytes()),
            );
            if win32_error.is_err() {
                RegCloseKey(inproc_key).0;
                RegCloseKey(prog_id_key).0;
                RegCloseKey(clsid_key).0;
                return Err(win32_error.into());
            }

            RegCloseKey(inproc_key).0;
            RegCloseKey(prog_id_key).0;
            RegCloseKey(clsid_key).0;

            Ok(())
        }
    }

    match register_com_object() {
        Ok(_) => S_OK,
        Err(_) => E_FAIL,
    }
}

#[no_mangle]
pub extern "system" fn DllUnregisterServer() -> HRESULT {
    fn unregister_com_object() -> windows::core::Result<()> {
        unsafe {
            // 1. Remove ProgID
            let c_subkey = CString::new(ProgID).unwrap();
            let win32_error = RegDeleteKeyExA(
                HKEY_CLASSES_ROOT,
                PCSTR::from_raw(c_subkey.as_ptr() as *const u8),
                0x0200 | 0x0100,
                None
            );
            if win32_error.is_err() {
                return Err(win32_error.into());
            }

            // 2. Remove CLSID
            let clsid_string = format!("CLSID\\{{{:?}}}", CLSID_Calculator);
            let c_clsid_string = CString::new(clsid_string).unwrap();
            let win32_error = RegDeleteKeyExA(
                HKEY_CLASSES_ROOT,
                PCSTR::from_raw(c_clsid_string.as_ptr() as *const u8),
                0x0200 | 0x0100,
                None
            );
            if win32_error.is_err() {
                return Err(win32_error.into());
            }

            Ok(())
        }
    }
    match unregister_com_object() {
        Ok(_) => S_OK,
        Err(_) => E_FAIL,
    }
}

#[no_mangle]
pub extern "system" fn DllCanUnloadNow() -> HRESULT {
    // Add proper reference counting if needed
    S_FALSE
}