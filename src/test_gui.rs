// Minimal test: does egui 0.33 render at all?
fn main() -> eframe::Result {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([400.0, 200.0])
            .with_always_on_top()
            .with_title("Test"),
        ..Default::default()
    };
    eframe::run_native("Test", options, Box::new(|_| Ok(Box::new(TestApp))))
}

struct TestApp;

impl eframe::App for TestApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("Hello from egui 0.33!");
            ui.label("If you can see this, rendering works.");
        });
    }
}
