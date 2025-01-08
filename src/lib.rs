#![allow(non_snake_case, non_upper_case_globals, non_camel_case_types)]
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        System::{
            Com::*,
            LibraryLoader::*,
            Registry::*
        }
    },
};
use std::sync::Mutex;
mod util;

// Define GUID for interface and class
pub const IID_ICalculator: GUID = GUID::from_u128(0xac39fd20_f263_49f7_9374_04c4a384c8aa);
pub const CLSID_Calculator: GUID = GUID::from_u128(0x8e7e4b8a_e909_4a03_acdc_4fb03c8c2182);
pub const ProgID: &str = "Calculator.Basic";

// Define the initialization settings structure
#[derive(Debug, Clone)]
struct CalculatorSettings {
    precision: i32,
    mode: String, // only for String type handling, not used!
}

// Define interface for COM object
#[interface("AC39FD20-F263-49F7-9374-04C4A384C8AA")]
pub unsafe trait ICalculator: IUnknown {
    fn Add(&self, a: f32, b: f32) -> core::result::Result<f32, Error>;
    fn Subtract(&self, a: f32, b: f32) -> core::result::Result<f32, Error>;
    fn Multiply(&self, a: f32, b: f32) -> core::result::Result<f32, Error>;
    fn Divide(&self, a: f32, b: f32) -> core::result::Result<f32, Error>;
    fn SetPrecision(&self, precision: i32) -> core::result::Result<(), Error>;
    fn GetPrecision(&self) -> core::result::Result<i32, Error>;
    fn SetMode(&self, mode: &str) -> core::result::Result<(), Error>;
    fn GetMode(&self) -> core::result::Result<BSTR, Error>;
}

// Define COM object
#[implement(ICalculator)]
pub struct Calculator {
    settings: Mutex<CalculatorSettings>,
}

impl Calculator {
    fn new(precision: i32, mode: &str) -> windows::core::Result<Self> {
        // Validate constructor parameters
        if precision < 0 || precision > 15 {
            return Err(Error::new(E_INVALIDARG, "Precision must be between 0 and 15"));
        }

        if !["standard", "scientific"].contains(&mode.to_lowercase().as_str()) {
            return Err(Error::new(E_INVALIDARG, "Mode must be 'standard' or 'scientific'"));
        }

        Ok(Self {
            settings: Mutex::new(CalculatorSettings {
                precision,
                mode: mode.to_string(),
            }),
        })
    }

    // Round helper
    fn round_value(&self, value: f32) -> f32 {
        let precision = self.settings.lock().unwrap().precision;
        let factor = 10.0f32.powi(precision);
        (value * factor).round() / factor
    }
}

impl ICalculator_Impl for Calculator_Impl {
    unsafe fn Add(&self, a: f32, b: f32) -> windows::core::Result<f32> {
        let result = a + b;
        Ok(self.round_value(result))
    }

    unsafe fn Subtract(&self, a: f32, b: f32) -> windows::core::Result<f32> {
        let result = a - b;
        Ok(self.round_value(result))
    }

    unsafe fn Multiply(&self, a: f32, b: f32) -> windows::core::Result<f32> {
        let result = a * b;
        Ok(self.round_value(result))
    }

    unsafe fn Divide(&self, a: f32, b: f32) -> windows::core::Result<f32> {
        if b == 0.0 {
            return Err(Error::new(E_INVALIDARG, "Division by zero"));
        }
        let result = a / b;
        Ok(self.round_value(result))
    }

    unsafe fn SetPrecision(&self, precision: i32) -> windows::core::Result<()> {
        match self.settings.lock() {
            Ok(mut settings) => {
                settings.precision = precision;
                Ok(())
            }
            Err(_) => Err(Error::new(E_FAIL, "Mutex poisoning"))
        }
    }

    unsafe fn GetPrecision(&self) -> windows::core::Result<i32> {
        Ok(self.settings.lock().unwrap().precision)
    }

    unsafe fn SetMode(&self, mode: &str) -> windows::core::Result<()> {
        match self.settings.lock() {
            Ok(mut settings) => {
                settings.mode = mode.to_string();
                Ok(())
            }
            Err(_) => Err(Error::new(E_FAIL, "Mutex poisoning"))
        }
    }

    unsafe fn GetMode(&self) -> windows::core::Result<BSTR> {
        Ok(BSTR::from(self.settings.lock().unwrap().mode.as_str()))
    }
}

// ClassFactory for Calculator
#[implement(IClassFactory)]
pub struct CalculatorFactory;

impl CalculatorFactory {
    fn new() -> Self {
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

    let factory: CalculatorFactory = CalculatorFactory::new();
    let unknown: windows::core::IUnknown = factory.into(); // tady by melo byt spis into ICalculator?
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
            let win32_error = RegCreateKeyA(
                HKEY_CLASSES_ROOT,
                PCSTR(ProgID.as_bytes().as_ptr()),
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
                PCSTR("CLSID".as_bytes().as_ptr()),
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
            let clsid_key_path = format!("CLSID\\{}", clsid_string);
            let mut clsid_key = HKEY::default();
            let win32_error = RegCreateKeyA(
                HKEY_CLASSES_ROOT,
                PCSTR(clsid_key_path.as_bytes().as_ptr()),
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
                PCSTR("ProgID".as_bytes().as_ptr()),
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
                PCSTR("InprocServer32".as_bytes().as_ptr()),
                &mut inproc_key
            );
            if win32_error.is_err() {
                RegCloseKey(prog_id_key).0;
                RegCloseKey(clsid_key).0;
                return Err(win32_error.into());
            }

            let path = util::get_this_module_path()?;
            let win32_error = RegSetValueExA(
                inproc_key,
                None,
                None,
                REG_SZ,
                Some(path.align_to::<u8>().1),
            );
            if win32_error.is_err() {
                RegCloseKey(inproc_key).0;
                RegCloseKey(prog_id_key).0;
                RegCloseKey(clsid_key).0;
                return Err(win32_error.into());
            }

            let win32_error = RegSetValueExA(
                inproc_key,
                PCSTR("ThreadingModel".as_bytes().as_ptr()),
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
            let win32_error = RegDeleteKeyA(
                HKEY_CLASSES_ROOT,
                PCSTR(ProgID.as_bytes().as_ptr()),
            );
            if win32_error.is_err() {
                return Err(win32_error.into());
            }

            // 2. Remove CLSID
            let clsid_key = format!("CLSID\\{{{:?}}}", CLSID_Calculator);
            let win32_error = RegDeleteKeyA(
                HKEY_CLASSES_ROOT,
                PCSTR(clsid_key.as_bytes().as_ptr()),
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test() {
        // Inicializujeme COM knihovnu
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).unwrap() };

        // Zavoláme CoCreateInstance pro získání COM objektu
        let com_object: windows::core::IUnknown = unsafe { CoCreateInstance(&CLSID_Calculator, None, CLSCTX_ALL).unwrap() };

        // Můžeme použít `cast` pro získání rozhraní (např. IUnknown nebo jiného rozhraní)
        let calc: ICalculator = com_object.cast::<ICalculator>().unwrap();

        // Práce s COM objektem, volání metod, apod.
        println!("COM object obtained: {:?}", calc);
        println!("COM object precision: {:?}", unsafe { calc.GetPrecision() });

        // Uvolníme COM knihovnu
        unsafe { CoUninitialize() };

    }
}