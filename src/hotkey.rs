use global_hotkey::{GlobalHotKeyEvent, GlobalHotKeyManager, HotKeyState};
use keyboard_types::{Code, Modifiers};
use std::sync::mpsc::{channel, Receiver};

pub struct HotkeyHandler {
    #[allow(dead_code)]
    pub manager: GlobalHotKeyManager,
    pub rx: Receiver<()>,
}

pub fn setup_hotkey() -> HotkeyHandler {
    let manager = GlobalHotKeyManager::new().expect("创建热键管理器失败");

    let hotkey = global_hotkey::hotkey::HotKey::new(
        Some(Modifiers::ALT),
        Code::KeyZ,
    );
    manager.register(hotkey).expect("注册热键失败");

    let (tx, rx) = channel::<()>();

    std::thread::spawn(move || {
        loop {
            if let Ok(event) = GlobalHotKeyEvent::receiver().try_recv() {
                if event.state == HotKeyState::Pressed {
                    let _ = tx.send(());
                }
            }
            std::thread::sleep(std::time::Duration::from_millis(50));
        }
    });

    HotkeyHandler { manager, rx }
}
