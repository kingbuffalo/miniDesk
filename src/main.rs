mod app;
mod data;
mod hotkey;
mod icon;
mod single_instance;
mod tray;

use app::MiniQApp;
use single_instance::SingleInstance;

fn main() {
    // 单实例检查
    let _instance = match SingleInstance::new("MiniQDesktop_Rust_v1") {
        Some(i) => i,
        None => {
            eprintln!("MiniQDesk 已经在运行");
            std::process::exit(0);
        }
    };

    // 初始化 COM（部分 Win32 API 需要）
    unsafe {
        let _ = windows::Win32::System::Com::CoInitializeEx(
            None,
            windows::Win32::System::Com::COINIT_APARTMENTTHREADED,
        );
    }

    // 托盘
    let (tray, show_id, exit_id) = tray::setup_tray();

    // 全局热键
    let hotkey = hotkey::setup_hotkey();

    // eframe 配置
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([1000.0, 400.0])
            .with_always_on_top()
            .with_decorations(true)
            .with_transparent(true)
            .with_taskbar(true),
        ..Default::default()
    };

    let result = eframe::run_native(
        "MiniQ Desk",
        options,
        Box::new(|cc| {
            // 设置半透明暗色主题
            let mut visuals = egui::Visuals::dark();
            visuals.window_fill = egui::Color32::from_rgba_premultiplied(30, 30, 30, 220);
            visuals.panel_fill = egui::Color32::from_rgba_premultiplied(30, 30, 30, 220);
            visuals.widgets.inactive.weak_bg_fill = egui::Color32::from_gray(45);
            visuals.widgets.hovered.weak_bg_fill = egui::Color32::from_gray(60);
            cc.egui_ctx.set_visuals(visuals);

            Ok(Box::new(MiniQApp::new(cc, hotkey.rx, tray, show_id, exit_id)))
        }),
    );

    if let Err(e) = result {
        eprintln!("运行错误: {}", e);
    }
}
