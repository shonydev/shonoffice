# Rust docx PoC (fase inicial, legado)

Este código es la **fase inicial** de ShonOffice: una GUI en Rust
(`egui`/`eframe`) que abre un `.docx` y muestra su texto plano, usando
`docx-rs` para leer el archivo.

## Por qué está acá y no en `src/`

A partir del [nuevo rumbo arquitectónico](../../README.md#nuevo-rumbo-arquitectónico-c-por-defecto-rust-por-necesidad),
**C#/.NET es la tecnología principal de ShonOffice** y Rust deja de ser
una pieza obligatoria del proyecto. Esta carpeta:

- **No forma parte del build principal** (no está referenciada desde
  `ShonOffice.sln` ni desde ningún proyecto C#).
- **Ya está superada funcionalmente** por `ShonOffice.UI` (Avalonia) +
  `ShonOffice.Infra.OpenXml`, que no solo muestran el texto de un
  `.docx` sino su formato real (negrita, color, tablas, alineación,
  encabezados).
- Se conserva únicamente **como referencia histórica/PoC**: documenta la
  idea original de usar Rust para todo el procesamiento de documentos,
  antes de decidir que Open XML SDK (.NET) es, en la práctica, muchísimo
  más completo para `.docx`/`.xlsx`/`.pptx`.

## Cómo compilarlo (opcional)

Sigue siendo un crate de Cargo autocontenido:

```bash
cd legacy/rust-docx-poc
cargo run
```

## ¿Cuándo volvería a ser relevante Rust?

Según el nuevo rumbo, Rust podría reaparecer en el futuro, pero **no
acá**: como una implementación concreta de `IPdfEngine` (puerto del
Domain en C#), solo si se demuestra que hace falta para PDF (parsing,
rendering, procesamiento pesado) y siempre detrás de esa interfaz. Ver
"Próximos pasos sugeridos" en el README principal.
