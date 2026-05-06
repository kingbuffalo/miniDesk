use crate::data::{open_path, reveal_in_explorer, ShortcutData};
use crate::icon::extract_icon;
use eframe::Frame;
use egui::{Context, Id, Pos2, Vec2};
use std::collections::HashMap;
use std::sync::mpsc::Receiver;

pub struct MiniQApp {
    pub data: ShortcutData,
    icon_cache: HashMap<String, egui::TextureHandle>,
    expanded_group: Option<String>,

    // 热键通道
    hotkey_rx: Receiver<()>,

    // 托盘
    #[allow(dead_code)]
    _tray: tray_icon::TrayIcon,
    show_menu_id: tray_icon::menu::MenuId,
    exit_menu_id: tray_icon::menu::MenuId,

    // 窗口状态
    hidden: bool,
    should_exit: bool,

    // 新建分组
    new_group_input: String,
    show_new_group_input: bool,

    // 添加文件/文件夹时的分组选择
    pending_add: Option<(String, String)>, // (path, name)
    add_target_group: String,

    // 删除确认
    pending_delete: Option<(String, usize)>, // (group, index)
}

impl MiniQApp {
    pub fn new(
        _cc: &eframe::CreationContext<'_>,
        hotkey_rx: Receiver<()>,
        tray: tray_icon::TrayIcon,
        show_menu_id: tray_icon::menu::MenuId,
        exit_menu_id: tray_icon::menu::MenuId,
    ) -> Self {
        let data = ShortcutData::load();

        Self {
            data,
            icon_cache: HashMap::new(),
            expanded_group: None,
            hotkey_rx,
            _tray: tray,
            show_menu_id,
            exit_menu_id,
            hidden: false,
            should_exit: false,
            new_group_input: String::new(),
            show_new_group_input: false,
            pending_add: None,
            add_target_group: String::new(),
            pending_delete: None,
        }
    }

    fn get_icon(&mut self, ctx: &Context, path: &str) -> Option<egui::TextureHandle> {
        if let Some(tex) = self.icon_cache.get(path) {
            return Some(tex.clone());
        }
        let img = extract_icon(path, 48)?;
        let width = img.width() as usize;
        let height = img.height() as usize;
        let pixels = img.into_raw();
        let color_image = egui::ColorImage::from_rgba_unmultiplied([width, height], &pixels);
        let texture = ctx.load_texture(path, color_image, egui::TextureOptions::LINEAR);
        self.icon_cache.insert(path.to_string(), texture.clone());
        Some(texture)
    }

    fn truncate_name(name: &str, max_chars: usize) -> String {
        let chars: Vec<char> = name.chars().collect();
        if chars.len() <= max_chars {
            name.to_string()
        } else {
            let first: String = chars[..max_chars].iter().collect();
            let second: String = chars[max_chars..chars.len().min(max_chars * 2)]
                .iter()
                .collect();
            if chars.len() > max_chars * 2 {
                format!("{}\n{}...", first, second)
            } else {
                format!("{}\n{}", first, second)
            }
        }
    }

    fn handle_external_events(&mut self, ctx: &Context) {
        // 托盘菜单事件
        while let Ok(event) = tray_icon::menu::MenuEvent::receiver().try_recv() {
            if event.id == self.show_menu_id {
                self.hidden = false;
            } else if event.id == self.exit_menu_id {
                self.should_exit = true;
                ctx.send_viewport_cmd(egui::ViewportCommand::Close);
            }
        }

        // 热键事件
        while let Ok(()) = self.hotkey_rx.try_recv() {
            self.hidden = !self.hidden;
        }
    }

    fn update_window_state(&self, ctx: &Context) {
        let screen_rect = ctx.input(|i| i.screen_rect());
        if self.hidden {
            // 完全隐藏窗口，不再留 MQ 小条（避免位置/标题栏异常问题）
            ctx.send_viewport_cmd(egui::ViewportCommand::Visible(false));
        } else {
            ctx.send_viewport_cmd(egui::ViewportCommand::Visible(true));
            ctx.send_viewport_cmd(egui::ViewportCommand::WindowLevel(
                egui::WindowLevel::AlwaysOnTop,
            ));
            ctx.send_viewport_cmd(egui::ViewportCommand::InnerSize(Vec2::new(1000.0, 400.0)));
            let y = (screen_rect.max.y - 400.0 - 60.0).max(0.0);
            ctx.send_viewport_cmd(egui::ViewportCommand::OuterPosition(Pos2::new(0.0, y)));
        }
    }

    fn render_hidden_ui(&mut self, ctx: &Context) {
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.vertical_centered(|ui| {
                ui.add_space(20.0);
                let btn = ui.add_sized(
                    [36.0, 80.0],
                    egui::Button::new(egui::RichText::new("MQ").size(16.0).strong()),
                );
                if btn.clicked() {
                    self.hidden = false;
                }
            });
        });
    }

    fn render_main_ui(&mut self, ctx: &Context) {
        // 底部固定控制栏
        egui::TopBottomPanel::bottom("controls")
            .exact_height(50.0)
            .show(ctx, |ui| {
                self.render_controls(ui);
            });

        // 中间滚动内容区
        egui::CentralPanel::default().show(ctx, |ui| {
            egui::ScrollArea::vertical().show(ui, |ui| {
                let groups: Vec<String> = self.data.groups.keys().cloned().collect();
                for group_name in groups {
                    self.render_group(ui, &group_name);
                }
            });
        });

        // 分组选择弹窗
        self.render_group_select_dialog(ctx);

        // 删除确认弹窗
        self.render_delete_confirm(ctx);
    }

    fn render_group(&mut self, ui: &mut egui::Ui, group_name: &str) {
        let is_expanded = self.expanded_group.as_deref() == Some(group_name);
        let id = Id::new(format!("group_{}", group_name));

        let shortcuts = self.data.groups.get(group_name).cloned().unwrap_or_default();
        let resp = egui::CollapsingHeader::new(
            egui::RichText::new(group_name).size(16.0).strong(),
        )
        .id_source(id)
        .open(Some(is_expanded))
        .show(ui, |ui| {
            if shortcuts.is_empty() {
                ui.label("（空分组）");
            } else {
                let max_cols = 4;
                egui::Grid::new(format!("grid_{}", group_name))
                    .spacing([16.0, 12.0])
                    .show(ui, |ui| {
                        for (i, shortcut) in shortcuts.iter().enumerate() {
                            if i > 0 && i % max_cols == 0 {
                                ui.end_row();
                            }
                            self.render_shortcut(ui, shortcut, group_name, i);
                        }
                    });
            }
        });

        if resp.header_response.clicked() {
            if is_expanded {
                self.expanded_group = None;
            } else {
                self.expanded_group = Some(group_name.to_string());
            }
        }

        // 右键删除分组
        resp.header_response.context_menu(|ui| {
            if ui.button("Delete Group").clicked() {
                self.data.remove_group(group_name);
                if self.expanded_group.as_deref() == Some(group_name) {
                    self.expanded_group = None;
                }
                ui.close_menu();
            }
        });
    }

    fn render_shortcut(
        &mut self,
        ui: &mut egui::Ui,
        shortcut: &crate::data::Shortcut,
        group: &str,
        index: usize,
    ) {
        let size = Vec2::new(80.0, 90.0);
        let (rect, response) = ui.allocate_exact_size(size, egui::Sense::click());

        let bg_color = if response.hovered() {
            egui::Color32::from_gray(60)
        } else {
            egui::Color32::from_gray(40)
        };

        ui.painter().rect_filled(rect, 6.0, bg_color);

        if response.is_pointer_button_down_on() {
            ui.painter()
                .rect_stroke(rect, 6.0, (1.0, egui::Color32::LIGHT_BLUE));
        }

        ui.allocate_ui_at_rect(rect, |ui| {
            ui.vertical_centered(|ui| {
                ui.add_space(6.0);
                // 图标
                if let Some(tex) = self.get_icon(ui.ctx(), &shortcut.path) {
                    ui.image((tex.id(), Vec2::new(40.0, 40.0)));
                } else {
                    ui.allocate_space(Vec2::new(40.0, 40.0));
                }
                ui.add_space(4.0);
                // 文字
                let text = Self::truncate_name(&shortcut.name, 8);
                ui.label(egui::RichText::new(text).size(11.0));
            });
        });

        // 右键菜单
        response.context_menu(|ui| {
            if ui.button("Open").clicked() {
                open_path(&shortcut.path);
                ui.close_menu();
            }
            if ui.button("Reveal").clicked() {
                reveal_in_explorer(&shortcut.path);
                ui.close_menu();
            }
            if ui.button("Delete").clicked() {
                self.pending_delete = Some((group.to_string(), index));
                ui.close_menu();
            }
        });

        // 左键打开
        if response.clicked() {
            open_path(&shortcut.path);
        }
    }

    fn render_controls(&mut self, ui: &mut egui::Ui) {
        ui.horizontal(|ui| {
            if ui.button("+ File").clicked() {
                if let Some(path) = rfd::FileDialog::new().pick_file() {
                    let name = path
                        .file_stem()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_string();
                    let path_str = path.to_string_lossy().to_string();
                    if self.data.groups.is_empty() {
                        self.data.add_shortcut("默认分组".to_string(), name, path_str);
                    } else {
                        self.pending_add = Some((path_str, name));
                        self.add_target_group = self.data.groups.keys().next().unwrap().clone();
                    }
                }
            }

            if ui.button("++ Folder").clicked() {
                if let Some(path) = rfd::FileDialog::new().pick_folder() {
                    let name = path
                        .file_name()
                        .unwrap_or_default()
                        .to_string_lossy()
                        .to_string();
                    let path_str = path.to_string_lossy().to_string();
                    if self.data.groups.is_empty() {
                        self.data.add_shortcut("默认分组".to_string(), name, path_str);
                    } else {
                        self.pending_add = Some((path_str, name));
                        self.add_target_group = self.data.groups.keys().next().unwrap().clone();
                    }
                }
            }

            ui.separator();

            if !self.show_new_group_input {
                if ui.button("[+] New Group").clicked() {
                    self.show_new_group_input = true;
                    self.new_group_input.clear();
                }
            } else {
                ui.horizontal(|ui| {
                    ui.text_edit_singleline(&mut self.new_group_input);
                    if ui.button("确认").clicked() && !self.new_group_input.is_empty() {
                        self.data.add_group(self.new_group_input.clone());
                        self.new_group_input.clear();
                        self.show_new_group_input = false;
                    }
                    if ui.button("Cancel").clicked() {
                        self.show_new_group_input = false;
                    }
                });
            }
        });
    }

    fn render_group_select_dialog(&mut self, ctx: &Context) {
        if self.pending_add.is_none() {
            return;
        }

        let mut should_close = false;
        let mut do_add: Option<(String, String, String)> = None; // (group, name, path)

        egui::Window::new("Select Group")
            .collapsible(false)
            .resizable(false)
            .show(ctx, |ui| {
                ui.label("Select target group:");
                ui.add_space(8.0);

                for group in self.data.groups.keys() {
                    if ui.selectable_label(self.add_target_group == *group, group).clicked() {
                        self.add_target_group = group.clone();
                    }
                }

                ui.separator();

                ui.horizontal(|ui| {
                    if ui.button("OK").clicked() {
                        if let Some((path, name)) = self.pending_add.take() {
                            do_add = Some((self.add_target_group.clone(), name, path));
                        }
                    }
                    if ui.button("Cancel").clicked() {
                        should_close = true;
                    }
                });
            });

        if should_close {
            self.pending_add = None;
        }
        if let Some((group, name, path)) = do_add {
            self.data.add_shortcut(group, name, path);
        }
    }

    fn render_delete_confirm(&mut self, ctx: &Context) {
        if self.pending_delete.is_none() {
            return;
        }

        let mut confirmed = false;
        let mut cancelled = false;

        egui::Window::new("Confirm Delete")
            .collapsible(false)
            .resizable(false)
            .show(ctx, |ui| {
                ui.label("Delete this shortcut?");
                ui.add_space(8.0);
                ui.horizontal(|ui| {
                    if ui.button("OK").clicked() {
                        confirmed = true;
                    }
                    if ui.button("取消").clicked() {
                        cancelled = true;
                    }
                });
            });

        if confirmed {
            if let Some((group, index)) = self.pending_delete.take() {
                self.data.remove_shortcut(&group, index);
            }
        } else if cancelled {
            self.pending_delete = None;
        }
    }
}

impl eframe::App for MiniQApp {
    fn update(&mut self, ctx: &Context, _frame: &mut Frame) {
        // 拦截窗口关闭事件，转为隐藏到托盘
        if ctx.input(|i| i.viewport().close_requested()) {
            ctx.send_viewport_cmd(egui::ViewportCommand::CancelClose);
            self.hidden = true;
        }

        self.handle_external_events(ctx);
        self.update_window_state(ctx);

        if self.hidden {
            self.render_hidden_ui(ctx);
        } else {
            self.render_main_ui(ctx);
        }
    }
}
