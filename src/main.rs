mod bridge;

use bridge::{CursorContext, TextBridge};
use std::time::{Duration, Instant};

// --- Bridge manager: picks the best available bridge ---

struct BridgeManager {
    bridges: Vec<Box<dyn TextBridge>>,
    last_check: Instant,
}

impl BridgeManager {
    fn new() -> Self {
        let mut bridges: Vec<Box<dyn TextBridge>> = Vec::new();

        #[cfg(target_os = "windows")]
        {
            if let Some(word) = bridge::word_com::WordComBridge::try_connect() {
                println!("Word COM bridge connected");
                bridges.push(Box::new(word));
            }
            bridges.push(Box::new(bridge::accessibility_win::AccessibilityBridge::new()));
        }

        BridgeManager {
            bridges,
            last_check: Instant::now(),
        }
    }

    fn read_context(&mut self) -> Option<CursorContext> {
        #[cfg(target_os = "windows")]
        if self.last_check.elapsed() > Duration::from_secs(5) {
            self.last_check = Instant::now();
            let has_word = self.bridges.iter().any(|b| b.name() == "Word COM");
            if !has_word {
                if let Some(word) = bridge::word_com::WordComBridge::try_connect() {
                    println!("Word COM bridge connected (late)");
                    self.bridges.insert(0, Box::new(word));
                }
            }
        }

        for bridge in &self.bridges {
            if bridge.is_available() {
                if let Some(ctx) = bridge.read_context() {
                    return Some(ctx);
                }
            }
        }
        None
    }

    fn active_bridge_name(&self) -> &str {
        for bridge in &self.bridges {
            if bridge.is_available() {
                return bridge.name();
            }
        }
        "none"
    }

    #[allow(dead_code)]
    fn replace_word(&self, new_text: &str) -> bool {
        for bridge in &self.bridges {
            if bridge.is_available() {
                return bridge.replace_word(new_text);
            }
        }
        false
    }
}

// --- egui app ---

struct ContextApp {
    manager: BridgeManager,
    context: CursorContext,
    last_poll: Instant,
    poll_interval: Duration,
    follow_cursor: bool,
    last_caret_pos: Option<(i32, i32)>,
}

impl ContextApp {
    fn new() -> Self {
        #[cfg(target_os = "windows")]
        unsafe {
            use windows::Win32::System::Com::*;
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok();
        }

        ContextApp {
            manager: BridgeManager::new(),
            context: CursorContext::default(),
            last_poll: Instant::now(),
            poll_interval: Duration::from_millis(300),
            follow_cursor: true,
            last_caret_pos: None,
        }
    }
}

fn get_screen_size() -> (f32, f32) {
    #[cfg(target_os = "windows")]
    unsafe {
        use windows::Win32::UI::WindowsAndMessaging::*;
        let w = GetSystemMetrics(SM_CXSCREEN);
        let h = GetSystemMetrics(SM_CYSCREEN);
        return (w as f32, h as f32);
    }
    #[allow(unreachable_code)]
    (1920.0, 1080.0)
}

impl eframe::App for ContextApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        // Poll for new context
        if self.last_poll.elapsed() >= self.poll_interval {
            self.last_poll = Instant::now();
            if let Some(new_ctx) = self.manager.read_context() {
                if new_ctx.caret_pos.is_some() {
                    self.last_caret_pos = new_ctx.caret_pos;
                }
                self.context = new_ctx;
            }
        }

        // Follow cursor: move window + hide title bar
        const WIN_W: f32 = 420.0;
        const WIN_H: f32 = 110.0;

        // Always borderless — our header row is the drag handle
        ctx.send_viewport_cmd(egui::ViewportCommand::Decorations(false));

        if self.follow_cursor {
            ctx.send_viewport_cmd(egui::ViewportCommand::InnerSize(egui::vec2(WIN_W, WIN_H)));
            if let Some((x, y)) = self.last_caret_pos {
                let (screen_w, screen_h) = get_screen_size();
                let pos_y = if (y as f32 + WIN_H) > screen_h {
                    y as f32 - WIN_H - 30.0
                } else {
                    y as f32
                };
                let pos_x = (x as f32).min(screen_w - WIN_W).max(0.0);

                ctx.send_viewport_cmd(egui::ViewportCommand::OuterPosition(
                    egui::pos2(pos_x, pos_y),
                ));
            }
        }

        ctx.request_repaint_after(Duration::from_millis(100));

        // Style
        let panel_frame = egui::Frame::new()
            .fill(egui::Color32::from_rgb(255, 255, 235))
            .stroke(egui::Stroke::new(1.0, egui::Color32::from_rgb(180, 170, 140)))
            .inner_margin(8.0);

        egui::CentralPanel::default().frame(panel_frame).show(ctx, |ui| {
            // Top row: checkbox + drag area + bridge name
            ui.horizontal(|ui| {
                ui.checkbox(&mut self.follow_cursor,
                    egui::RichText::new("Følg cursor").size(12.0)
                        .color(egui::Color32::from_rgb(60, 60, 55))
                );

                // Drag handle area (empty space between checkbox and bridge name)
                let remaining = ui.available_rect_before_wrap();
                let drag_resp = ui.allocate_rect(remaining, egui::Sense::drag());
                if drag_resp.drag_started() && !self.follow_cursor {
                    ctx.send_viewport_cmd(egui::ViewportCommand::StartDrag);
                }

                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                    ui.label(
                        egui::RichText::new(self.manager.active_bridge_name())
                            .size(10.0)
                            .color(egui::Color32::from_rgb(160, 155, 140)),
                    );
                });
            });

            ui.separator();

            // Word
            if !self.context.word.is_empty() {
                ui.label(
                    egui::RichText::new(&self.context.word)
                        .strong()
                        .size(15.0)
                        .color(egui::Color32::from_rgb(0, 70, 160)),
                );
                ui.add_space(2.0);
            }

            // Sentence
            if !self.context.sentence.is_empty() {
                ui.label(
                    egui::RichText::new(&self.context.sentence)
                        .size(12.0)
                        .color(egui::Color32::from_rgb(50, 50, 50)),
                );
            }

            if self.context.word.is_empty() && self.context.sentence.is_empty() {
                ui.label(
                    egui::RichText::new("Flytt cursoren for aa se kontekst...")
                        .italics()
                        .size(11.0)
                        .color(egui::Color32::from_rgb(150, 150, 140)),
                );
            }
        });
    }
}

fn main() -> eframe::Result {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([400.0, 100.0])
            .with_always_on_top()
            .with_decorations(true)
            .with_title("NorskTale"),
        ..Default::default()
    };

    eframe::run_native(
        "NorskTale",
        options,
        Box::new(|_cc| Ok(Box::new(ContextApp::new()))),
    )
}
