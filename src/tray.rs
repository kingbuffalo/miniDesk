use tray_icon::{
    menu::{Menu, MenuItem, PredefinedMenuItem},
    TrayIconBuilder,
};

pub fn setup_tray() -> (tray_icon::TrayIcon, tray_icon::menu::MenuId, tray_icon::menu::MenuId) {
    let menu = Menu::new();
    let show_item = MenuItem::new("Show", true, None);
    let sep = PredefinedMenuItem::separator();
    let exit_item = MenuItem::new("Exit", true, None);

    let _ = menu.append(&show_item);
    let _ = menu.append(&sep);
    let _ = menu.append(&exit_item);

    let icon_bytes = include_bytes!("../icons/icon.png");
    let icon_image = image::load_from_memory(icon_bytes)
        .expect("加载托盘图标失败")
        .into_rgba8();
    let (width, height) = icon_image.dimensions();
    let icon = tray_icon::Icon::from_rgba(icon_image.into_raw(), width, height)
        .expect("创建托盘图标失败");

    let tray = TrayIconBuilder::new()
        .with_menu(Box::new(menu))
        .with_tooltip("MiniQ Desk")
        .with_icon(icon)
        .build()
        .expect("创建托盘图标失败");

    let show_id = show_item.id().clone();
    let exit_id = exit_item.id().clone();

    (tray, show_id, exit_id)
}
