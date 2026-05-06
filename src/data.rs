use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use std::path::PathBuf;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Shortcut {
    pub name: String,
    pub path: String,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct ShortcutData {
    #[serde(default)]
    pub groups: BTreeMap<String, Vec<Shortcut>>,
}

impl ShortcutData {
    pub fn load() -> Self {
        let path = Self::data_path();
        if path.exists() {
            if let Ok(content) = std::fs::read_to_string(&path) {
                if let Ok(data) = serde_json::from_str::<ShortcutData>(&content) {
                    return data;
                }
            }
        }
        Self::default()
    }

    pub fn save(&self) {
        let path = Self::data_path();
        if let Ok(json) = serde_json::to_string_pretty(self) {
            let _ = std::fs::write(&path, json);
        }
    }

    fn data_path() -> PathBuf {
        std::env::current_dir()
            .unwrap_or_else(|_| PathBuf::from("."))
            .join("shortcuts.json")
    }

    pub fn add_shortcut(&mut self, group: String, name: String, path: String) {
        self.groups
            .entry(group)
            .or_default()
            .push(Shortcut { name, path });
        self.save();
    }

    pub fn remove_shortcut(&mut self, group: &str, index: usize) {
        if let Some(list) = self.groups.get_mut(group) {
            if index < list.len() {
                list.remove(index);
                if list.is_empty() {
                    self.groups.remove(group);
                }
                self.save();
            }
        }
    }

    pub fn add_group(&mut self, name: String) {
        self.groups.entry(name).or_default();
        self.save();
    }

    pub fn remove_group(&mut self, name: &str) {
        self.groups.remove(name);
        self.save();
    }

    pub fn rename_group(&mut self, old: &str, new: String) {
        if let Some(items) = self.groups.remove(old) {
            self.groups.insert(new, items);
            self.save();
        }
    }
}

pub fn open_path(path: &str) {
    let _ = std::process::Command::new("cmd")
        .args(["/C", "start", "", path])
        .spawn();
}

pub fn reveal_in_explorer(path: &str) {
    let _ = std::process::Command::new("explorer")
        .args(["/select,", path])
        .spawn();
}
