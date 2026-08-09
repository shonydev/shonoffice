use docx_rs::*;

/// Lee un archivo .docx desde disco y devuelve su texto, un párrafo por línea.
pub fn read_docx_text(bytes: &[u8]) -> Result<Vec<String>, String> {
    let docx = read_docx(bytes).map_err(|e| format!("{:?}", e))?;

    let mut paragraphs = Vec::new();

    for child in docx.document.children.iter() {
        if let DocumentChild::Paragraph(paragraph) = child {
            let text = extract_paragraph_text(paragraph);
            if !text.trim().is_empty() {
                paragraphs.push(text);
            }
        }
    }

    Ok(paragraphs)
}

/// Extrae el texto plano de un parrafo, concatenando todos sus "runs".
fn extract_paragraph_text(paragraph: &Paragraph) -> String {
    let mut text = String::new();

    for child in paragraph.children.iter() {
        if let ParagraphChild::Run(run) = child {
            for run_child in run.children.iter() {
                if let RunChild::Text(t) = run_child {
                    text.push_str(&t.text);
                }
            }
        }
    }

    text
}
