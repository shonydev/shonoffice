use eframe::egui;
use shonoffice::read_docx_text;
use std::fs;
use std::path::PathBuf;

fn main() -> eframe::Result<()> {
    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default().with_inner_size([700.0, 550.0]),
        ..Default::default()
    };

    eframe::run_native(
        "ShonOffice",
        options,
        Box::new(|_cc| Ok(Box::new(ShonOfficeApp::default()))),
    )
}

#[derive(Default)]
struct ShonOfficeApp {
    file_path: Option<PathBuf>,
    paragraphs: Vec<String>,
    error: Option<String>,
}

impl ShonOfficeApp {
    fn open_file(&mut self) {
        self.error = None;

        let Some(path) = rfd::FileDialog::new()
            .add_filter("Word (*.docx)", &["docx"])
            .pick_file()
        else {
            return; // el usuario cancelo el dialogo
        };

        match fs::read(&path) {
            Ok(bytes) => match read_docx_text(&bytes) {
                Ok(paragraphs) => {
                    self.paragraphs = paragraphs;
                    self.file_path = Some(path);
                }
                Err(e) => {
                    self.error = Some(format!("No se pudo leer el documento: {}", e));
                    self.paragraphs.clear();
                }
            },
            Err(e) => {
                self.error = Some(format!("No se pudo abrir el archivo: {}", e));
                self.paragraphs.clear();
            }
        }
    }
}

impl eframe::App for ShonOfficeApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        egui::TopBottomPanel::top("top_bar").show(ctx, |ui| {
            ui.add_space(6.0);
            ui.horizontal(|ui| {
                if ui.button("📂 Abrir Word...").clicked() {
                    self.open_file();
                }

                if let Some(path) = &self.file_path {
                    ui.label(format!("{}", path.display()));
                }
            });
            ui.add_space(6.0);
        });

        egui::CentralPanel::default().show(ctx, |ui| {
            if let Some(err) = &self.error {
                ui.colored_label(egui::Color32::RED, err);
                return;
            }

            if self.paragraphs.is_empty() {
                ui.centered_and_justified(|ui| {
                    ui.label("Abre un archivo .docx para ver su contenido.");
                });
                return;
            }

            egui::ScrollArea::vertical().show(ui, |ui| {
                for p in &self.paragraphs {
                    ui.label(p);
                    ui.add_space(4.0);
                }
            });
        });
    }
}
