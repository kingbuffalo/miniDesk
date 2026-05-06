use windows::Win32::Foundation::{CloseHandle, HANDLE};
use windows::Win32::System::Threading::CreateMutexW;

pub struct SingleInstance {
    handle: HANDLE,
}

impl SingleInstance {
    pub fn new(name: &str) -> Option<Self> {
        let wide: Vec<u16> = name.encode_utf16().chain(Some(0)).collect();
        unsafe {
            let handle = CreateMutexW(None, true, windows::core::PCWSTR(wide.as_ptr()));
            match handle {
                Ok(h) => {
                    let err = windows::core::Error::from_win32().code().0;
                    if err == 183 {
                        // ERROR_ALREADY_EXISTS
                        let _ = CloseHandle(h);
                        return None;
                    }
                    Some(SingleInstance { handle: h })
                }
                Err(_) => None,
            }
        }
    }
}

impl Drop for SingleInstance {
    fn drop(&mut self) {
        unsafe {
            let _ = CloseHandle(self.handle);
        }
    }
}
