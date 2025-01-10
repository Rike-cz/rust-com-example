#![allow(non_snake_case, non_upper_case_globals, non_camel_case_types)]
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        System::Com::*,
    },
};
use std::sync::Mutex;
mod export;

// Define GUID for interface and class
pub const IID_ICalculator: GUID = GUID::from_u128(0xac39fd20_f263_49f7_9374_04c4a384c8aa);
pub const CLSID_Calculator: GUID = GUID::from_u128(0x8e7e4b8a_e909_4a03_acdc_4fb03c8c2182);
pub const ProgID: &str = "Calculator.Basic";

// Define the initialization settings structure
#[repr(C)]
#[derive(Debug, Clone)]
struct CalculatorSettings {
    precision: i32,
    mode: String, // only for String type handling, not used!
}

// Define interface for COM object
#[interface("AC39FD20-F263-49F7-9374-04C4A384C8AA")]
pub unsafe trait ICalculator: IUnknown {
    unsafe fn Add(&self, a: f32, b: f32) -> core::result::Result<f32, Error>;
    unsafe fn Add2(&self, a: f32, b: f32, res: *mut f32) -> HRESULT;
    unsafe fn Subtract(&self, a: f32, b: f32) -> core::result::Result<f32, Error>;
    unsafe fn Multiply(&self, a: f32, b: f32) -> core::result::Result<f32, Error>;
    unsafe fn Divide(&self, a: f32, b: f32) -> core::result::Result<f32, Error>;
    unsafe fn SetPrecision(&self, precision: i32) -> core::result::Result<(), Error>;
    unsafe fn GetPrecision(&self) -> core::result::Result<i32, Error>;
    unsafe fn SetMode(&self, mode: &str) -> core::result::Result<(), Error>;
    unsafe fn GetMode(&self) -> core::result::Result<BSTR, Error>;
}

// Define COM object
#[repr(C)]
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

// impl Calculator_Impl {
//     unsafe fn Initialize(&self, precision: i32, mode: HSTRING) -> windows::core::Result<Self> {
//         self.this.
//         this::new(precision, mode)
//     }
// }

impl ICalculator_Impl for Calculator_Impl {
    unsafe fn Add(&self, a: f32, b: f32) -> windows::core::Result<f32> {
        let result = a + b;
        Ok(self.round_value(result))
    }

    unsafe fn Add2(&self, a: f32, b: f32, res: *mut f32) -> HRESULT {
        *res = a + b;
        S_OK
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test() {
        // Initialize COM dll
        unsafe { CoInitializeEx(None, COINIT_MULTITHREADED).unwrap() };
        // match unsafe { CoInitializeEx(None, COINIT_MULTITHREADED) } {
        //     S_OK => (),
        //     S_FALSE => (),
        //     _ => (),
        // }

        // Get COM object
        let com_object: windows::core::IUnknown = unsafe { CoCreateInstance(&CLSID_Calculator, None, CLSCTX_INPROC_SERVER).unwrap() };

        // Casting
        let calc: ICalculator = com_object.cast::<ICalculator>().unwrap();

        // Try
        assert_eq!(unsafe { calc.GetPrecision().unwrap() }, 2);
        assert_eq!(unsafe { calc.Add(3.14, 6.86).unwrap() }, 10.0);

        let mut result: f32 = 0.0;
        unsafe { calc.Add2(3.14, 6.86, &mut result as *mut f32).unwrap() };
        assert_eq!(result, 10.0);
        assert_eq!(unsafe { calc.GetMode().unwrap() }, "standard");
        unsafe { calc.SetMode("scientific").unwrap() }
        assert_eq!(unsafe { calc.GetMode().unwrap() }, "scientific");

        // Release COM dll
        unsafe { CoUninitialize() };
        //println!("CoUninitialize");

        std::mem::forget(com_object);
        std::mem::forget(calc);

    }
}