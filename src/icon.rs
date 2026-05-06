use image::RgbaImage;
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;
use std::path::Path;
use windows::Win32::Foundation::HWND;
use windows::Win32::Graphics::Gdi::{
    DeleteObject, GetDC, GetDIBits, ReleaseDC, BITMAPINFO,
    BITMAPINFOHEADER, DIB_RGB_COLORS, HGDIOBJ,
};
use windows::Win32::UI::Shell::ExtractIconExW;
use windows::Win32::UI::WindowsAndMessaging::{DestroyIcon, HICON};
fn to_wide(s: &str) -> Vec<u16> {
    OsStr::new(s)
        .encode_wide()
        .chain(Some(0))
        .collect()
}

pub fn extract_icon(path: &str, size: u32) -> Option<RgbaImage> {
    let ext = Path::new(path)
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_lowercase();

    // 使用内置默认图标
    if ext == "exe" {
        return load_builtin_icon("icons/exe_icon.png", size);
    }
    if std::fs::metadata(path).ok()?.is_dir() {
        return load_builtin_icon("icons/default_folder_icon.png", size);
    }

    // 尝试从系统提取图标
    if let Some(img) = extract_icon_from_shell(path, size) {
        return Some(img);
    }

    None
}

fn load_builtin_icon(path: &str, size: u32) -> Option<RgbaImage> {
    let img = image::open(path).ok()?;
    Some(img.resize(size, size, image::imageops::FilterType::Lanczos3).to_rgba8())
}

fn extract_icon_from_shell(path: &str, size: u32) -> Option<RgbaImage> {
    let wide = to_wide(path);

    unsafe {
        let mut large_hicon = HICON(std::ptr::null_mut());
        let mut small_hicon = HICON(std::ptr::null_mut());
        let count = ExtractIconExW(
            windows::core::PCWSTR(wide.as_ptr()),
            0,
            Some(&mut large_hicon),
            Some(&mut small_hicon),
            1,
        );

        if count == 0 {
            return None;
        }

        let hicon = if !large_hicon.0.is_null() {
            large_hicon.0
        } else {
            small_hicon.0
        };

        if hicon.is_null() {
            return None;
        }

        let result = icon_to_rgba(hicon, size);

        let _ = DestroyIcon(windows::Win32::UI::WindowsAndMessaging::HICON(hicon));

        result
    }
}

unsafe fn icon_to_rgba(hicon: *mut std::ffi::c_void, size: u32) -> Option<RgbaImage> {
    // 获取图标信息
    let mut icon_info = std::mem::zeroed();
    if windows::Win32::UI::WindowsAndMessaging::GetIconInfo(
        windows::Win32::UI::WindowsAndMessaging::HICON(hicon),
        &mut icon_info,
    )
    .is_err()
    {
        return None;
    }

    let hdc_screen = GetDC(HWND(std::ptr::null_mut()));
    let hdc_mem = windows::Win32::Graphics::Gdi::CreateCompatibleDC(hdc_screen);

    let bmp_header = BITMAPINFOHEADER {
        biSize: std::mem::size_of::<BITMAPINFOHEADER>() as u32,
        biWidth: size as i32,
        biHeight: -(size as i32),
        biPlanes: 1,
        biBitCount: 32,
        biCompression: 0,
        biSizeImage: 0,
        biXPelsPerMeter: 0,
        biYPelsPerMeter: 0,
        biClrUsed: 0,
        biClrImportant: 0,
    };

    let mut bmi = BITMAPINFO {
        bmiHeader: bmp_header,
        bmiColors: [Default::default(); 1],
    };

    let hbitmap = windows::Win32::Graphics::Gdi::CreateDIBSection(
        hdc_mem,
        &bmi,
        DIB_RGB_COLORS,
        std::ptr::null_mut(),
        None,
        0,
    )
    .ok()?;

    let old = windows::Win32::Graphics::Gdi::SelectObject(hdc_mem, HGDIOBJ(hbitmap.0));

    // 绘制图标到内存 DC
    let _ = windows::Win32::UI::WindowsAndMessaging::DrawIconEx(
        hdc_mem,
        0,
        0,
        windows::Win32::UI::WindowsAndMessaging::HICON(hicon),
        size as i32,
        size as i32,
        0,
        None,
        windows::Win32::UI::WindowsAndMessaging::DI_NORMAL,
    );

    let mut pixels: Vec<u8> = vec![0; (size * size * 4) as usize];
    let _ = GetDIBits(
        hdc_mem,
        hbitmap,
        0,
        size,
        Some(pixels.as_mut_ptr() as *mut core::ffi::c_void),
        &mut bmi,
        DIB_RGB_COLORS,
    );

    windows::Win32::Graphics::Gdi::SelectObject(hdc_mem, old);
    let _ = DeleteObject(HGDIOBJ(hbitmap.0));
    let _ = windows::Win32::Graphics::Gdi::DeleteDC(hdc_mem);
    let _ = ReleaseDC(HWND(std::ptr::null_mut()), hdc_screen);

    // BGRA -> RGBA
    for chunk in pixels.chunks_exact_mut(4) {
        let b = chunk[0];
        let g = chunk[1];
        let r = chunk[2];
        let a = chunk[3];
        chunk[0] = r;
        chunk[1] = g;
        chunk[2] = b;
        chunk[3] = a;
    }

    RgbaImage::from_raw(size, size, pixels)
}
