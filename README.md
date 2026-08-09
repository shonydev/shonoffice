# ShonOffice

Suite de ofimática liviana para leer y (más adelante) editar archivos de
Microsoft Office (`.xlsx`, `.docx`, `.pptx`), con conversión PDF↔Word.

> **ShonOffice es, ante todo, una aplicación C#/.NET.** Rust ya no es una
> pieza obligatoria de la arquitectura — ver
> [Nuevo rumbo arquitectónico](#nuevo-rumbo-arquitectónico-c-por-defecto-rust-por-necesidad)
> más abajo.

## 🚀 Inicio rápido — ver un `.docx` con su formato real (no texto plano)

```bash
# Desde la raíz del repo (donde está ShonOffice.sln)
dotnet restore
dotnet build
dotnet run --project src/ShonOffice.UI
```

Esto abre la ventana **"ShonOffice"**. Click en **"📂 Open Word..."** 

**Requisitos:** [.NET SDK 8](https://dotnet.microsoft.com/download)
(`dotnet --version` debe mostrar `8.x`). La primera vez necesita conexión
a internet para restaurar los paquetes NuGet (Open XML SDK, Avalonia).


## Cómo ejecutar este proyecto

ShonOffice es una solución C#/.NET (`ShonOffice.sln`): un único comando la
compila y corre.

### UI en C# (Avalonia + Open XML SDK) — lee `.docx` con formato real

Esta es la UI para ver un `.docx` como se ve en Word de verdad: usa
`ShonOffice.Infra.OpenXml` (Open XML SDK) para leer el documento con su
formato real y `ShonOffice.UI` (Avalonia) para mostrarlo.

**Requisitos:** [.NET SDK 8](https://dotnet.microsoft.com/download) (`dotnet --version` debería mostrar `8.x`). La primera vez necesita conexión a internet para restaurar los paquetes NuGet (Open XML SDK, Avalonia).

```bash
# Desde la raíz del repo
dotnet restore
dotnet build
dotnet run --project src/ShonOffice.UI
```

Abre una ventana con un botón "Open Word..."

Al abrir un documento, esa barra superior (botón + ruta del archivo) se
oculta y el título de la ventana pasa a mostrar el nombre del archivo
abierto, igual que hace Word.

### PoC en Rust (legado, opcional)

`legacy/rust-docx-poc/` tiene la GUI en Rust/egui de la fase inicial del
proyecto, ya superada por la UI en C# de arriba. No forma parte del build
de `ShonOffice.sln` — se conserva solo como referencia histórica. Ver
[legacy/rust-docx-poc/README.md](legacy/rust-docx-poc/README.md) si
quieres compilarla igual.


## Nuevo rumbo arquitectónico: C# por defecto, Rust por necesidad

ShonOffice nació combinando dos tecnologías: **C#/.NET** para el
ecosistema de documentos de Office, y **Rust** para procesamiento
pesado. Después de revisar el estado del proyecto, decidimos ajustar ese
enfoque:

> **C# es la tecnología principal de ShonOffice. Rust se utiliza
> únicamente cuando existe una razón técnica concreta para hacerlo.**

Rust deja de ser una obligación arquitectónica y pasa a ser una
herramienta especializada, a usarse solo si se justifica.

### Qué significa esto en la práctica

- **UI, Domain, Application e Infraestructura de Office (`.docx`/`.xlsx`/`.pptx`) son C#/.NET**,
  aprovechando Avalonia y Open XML SDK — ver la sección de arquitectura
  hexagonal más abajo para el detalle de capas.
- **Rust ya no es la base del proyecto.** La GUI en Rust/egui de la fase
  inicial (`legacy/rust-docx-poc/`) queda como referencia histórica, no
  como un camino activo — está superada funcionalmente por
  `ShonOffice.UI` + `ShonOffice.Infra.OpenXml`.
- **PDF es un candidato para Rust, no una decisión tomada.** El puerto
  `IPdfEngine` (Domain) hoy apunta a `ShonOffice.Infra.Pdf`, una
  implementación 100% .NET (placeholder por ahora). Solo si se detecta
  un cuello de botella real — parsing, rendering, rasterización,
  procesamiento masivo — se evalúa una implementación en Rust detrás del
  mismo puerto (`ShonOffice.Infra.Native`, vía FFI), sin que el resto de
  ShonOffice necesite saber cuál de las dos está en uso:

  ```text
  IPdfEngine
      │
      ├── NotImplementedManagedPdfEngine   (ShonOffice.Infra.Pdf, C#, hoy)
      │
      └── RustPdfEngine                    (ShonOffice.Infra.Native, FFI, solo si se justifica)
  ```

- **Cuando (y si) Rust sea necesario**, debe estar aislado detrás de una
  interfaz del Domain, nunca filtrarse a la lógica de negocio ni a la UI.

### Criterio para decidir cuándo usar Rust

Ante cualquier funcionalidad nueva:

```text
¿C#/.NET resuelve correctamente el problema?
        │
       Sí → Usar C#
        │
       No
        ↓
¿Existe una alternativa .NET razonable?
        │
       Sí → Evaluarla
        │
       No
        ↓
¿Rust aporta una ventaja significativa (rendimiento, memoria,
procesamiento pesado, o una librería madura sin equivalente .NET)?
        │
       Sí → Usar Rust, detrás de una interfaz del Domain
       No → Usar C#
```

El orden de trabajo es *construir en C# → medir → identificar cuellos de
botella → evaluar alternativas .NET → recién ahí evaluar Rust*, nunca al
revés ("Rust primero, buscar después dónde usarlo"). Rust implica costos
adicionales (FFI, gestión de memoria entre runtimes, `unsafe`, builds
multiplataforma, CI/CD adicional, debugging de dos ecosistemas) que
deben justificarse con una ventaja técnica real, no asumirse de entrada.

**El objetivo no es usar dos lenguajes. El objetivo es construir la
mejor suite ofimática posible usando la tecnología adecuada para cada
problema.**

## Arquitectura de software para escalar: Hexagonal (Ports & Adapters)

Que C# sea la tecnología principal (y Rust, a lo sumo, un detalle de
infraestructura puntual) resuelve *qué tecnología hace qué*, pero no
resuelve *cómo se organiza el código para poder crecer* (sumar xlsx,
pptx, edición, conversión PDF, y eventualmente otra UI o un modo CLI)
sin que todo termine acoplado a Open XML SDK, a Avalonia, o a una
librería de PDF en particular. Para eso propongo **arquitectura
hexagonal (Ports & Adapters)**, con el dominio en el centro y todo lo
externo (librerías, UI, FFI) como piezas reemplazables alrededor.

### Por qué hexagonal y no otra cosa

- **Los formatos Office cambian de librería más seguido que la lógica de
  negocio.** Si hoy Open XML SDK resuelve todo pero mañana hace falta otra
  librería para un caso puntual (o un formato viejo, `.doc`/`.xls`), el
  dominio (modelo de documento, casos de uso) no debería enterarse. Eso
  solo es alcanzable si el dominio depende de **interfaces propias**
  (puertos) y no directamente de Open XML SDK.
- **Un futuro motor Rust sería, ni más ni menos, otro adaptador — nunca
  obligatorio.** Desde el punto de vista del dominio, "convertir PDF a
  Word" es un caso de uso que depende de un puerto `IPdfEngine`; que ese
  puerto lo implemente hoy una librería .NET, o mañana (solo si se
  justifica) un motor Rust vía FFI, es un detalle de infraestructura que
  no debería filtrarse a la lógica de negocio ni a la UI. Ver "Nuevo
  rumbo arquitectónico" más arriba.
- **Multiplica los puntos de entrada sin duplicar lógica.** UI, CLI, o
  incluso un futuro servicio/API son todos "adaptadores de entrada" que
  llaman a los mismos casos de uso de `ShonOffice.Application`.
- **Testeable sin abrir archivos reales.** Los casos de uso se pueden
  probar con adaptadores falsos (in-memory) para xlsx/docx/pptx/pdf, sin
  depender de Open XML SDK ni de ninguna librería externa.
- **Costo de entrada bajo para un proyecto de este tamaño.** Se descartó
  microservicios (no hay necesidad de escalar por separado ni de
  despliegue distribuido) y una arquitectura en capas estrictamente
  lineal (UI→Lógica→Datos) porque ahí Open XML SDK termina filtrándose
  hacia arriba tarde o temprano. Hexagonal da el aislamiento necesario
  con una curva de aprendizaje razonable.

### Capas y flujo de dependencias

```
                     ┌─────────────────────────────┐
   Adaptadores de    │   ShonOffice.UI (Avalonia)   │
   entrada           │   ShonOffice.Cli             │
                     └──────────────┬───────────────┘
                                    │ llama a
                     ┌──────────────▼───────────────┐
   Aplicación        │   ShonOffice.Application      │
   (casos de uso)    │   - OpenDocumentUseCase        │
                     │   - ConvertPdfToWordUseCase    │
                     │   - SaveDocumentUseCase        │
                     └──────────────┬───────────────┘
                                    │ depende de (interfaces)
                     ┌──────────────▼───────────────┐
   Dominio           │   ShonOffice.Domain           │
   (núcleo, sin      │   - OfficeDocument, Sheet, Slide...  │
   dependencias      │   - IDocxReader/Writer         │
   externas)         │   - IXlsxReader/Writer         │
                     │   - IPptxReader/Writer         │
                     │   - IPdfEngine                 │
                     └──────────────▲───────────────┘
                                    │ implementan (puertos)
              ┌─────────────────────┼─────────────────────┬───────────────────────┐
              │                     │                      │                       │
┌─────────────▼───────────┐ ┌───────▼────────────┐ ┌───────▼──────────┐ ┌──────────▼───────────┐
│ ShonOffice.Infra.OpenXml │ │ ShonOffice.Infra.Pdf │ │ (futuros adaptadores) │ │ ShonOffice.Infra.Native │
│ - implementa IDocxReader │ │ - implementa IPdfEngine │ │ - ej. exportar a ODF  │ │ (solo si Rust se       │
│   con Open Xml SDK       │ │   100% .NET (hoy)       │ │   sin tocar el resto  │ │  justifica para PDF)   │
└──────────────────────────┘ └─────────────────────┘ └───────────────────────┘ └─────────────────────────┘
```

`ShonOffice.Infra.Pdf` es la implementación "C# por defecto" de
`IPdfEngine`; `ShonOffice.Infra.Native` (Rust vía FFI) es un adaptador
alternativo del mismo puerto, a construir únicamente si se justifica
técnicamente — ver "Nuevo rumbo arquitectónico" más arriba. Ninguno de
los dos existe hoy con una implementación funcional: ambos son, por
ahora, puntos de extensión documentados.

Regla clave: **las flechas de dependencia siempre apuntan hacia adentro**.
El dominio no conoce Avalonia, ni Open XML SDK, ni Rust. La UI y los
adaptadores de infraestructura sí conocen al dominio (a través de sus
interfaces), nunca al revés. Esto es lo que permite, por ejemplo, cambiar
Avalonia por otra UI, o sumar un adaptador `.doc` legado, sin tocar los
casos de uso.

### Estructura de proyecto objetivo

```
shonoffice/
├── src/
│   ├── ShonOffice.Domain/          # C# — modelo de documento + puertos (interfaces)
│   ├── ShonOffice.Application/     # C# — casos de uso, orquesta los puertos
│   ├── ShonOffice.Infra.OpenXml/   # C# — implementa xlsx/docx/pptx con Open XML SDK
│   ├── ShonOffice.Infra.Pdf/       # C# — implementa IPdfEngine 100% .NET ("C# por defecto")
│   ├── ShonOffice.UI/              # C# + Avalonia — adaptador de entrada (MVVM)
│   ├── ShonOffice.Cli/             # C# — adaptador de entrada (terminal) (aún no creado)
│   └── ShonOffice.Infra.Native/    # C# — SOLO si Rust se justifica para PDF: bindings FFI a shonoffice-native (aún no creado)
├── tests/
│   ├── ShonOffice.Application.Tests/    # casos de uso con adaptadores falsos (in-memory)
│   └── ShonOffice.Infra.OpenXml.Tests/  # valida la resolución de estilos con un .docx construido en memoria
├── legacy/
│   └── rust-docx-poc/               # Rust — fase inicial, PoC histórico, no se compila junto al resto
└── README.md
```

`ShonOffice.Infra.Native` y su contraparte en Rust (`shonoffice-native`,
como `cdylib`) solo se crean si, siguiendo el criterio de "Nuevo rumbo
arquitectónico", se justifica técnicamente introducir Rust para PDF. No
son parte del plan por defecto.

En la UI (`ShonOffice.UI`), dentro de la arquitectura hexagonal, se usa
**MVVM** (el patrón nativo de Avalonia): las Views son el detalle visual,
los ViewModels llaman a los casos de uso de `ShonOffice.Application` y
exponen el estado a la vista. MVVM y hexagonal no compiten: MVVM ordena
*adentro* del adaptador de UI, hexagonal ordena la relación entre ese
adaptador y el resto del sistema.

# Estado actual

Seguimos avanzando según misiones atómicas:

- ✅ **Leer Word (.docx) con formato real** — `ShonOffice.Infra.OpenXml` implementa `IDocxReader` con Open XML SDK, resolviendo negrita, cursiva, subrayado, tamaño, color, fuente, alineación, sangría y nivel de encabezado a través de la cadena de estilos del documento (no solo formato directo) — y `ShonOffice.UI` (Avalonia) la muestra así, primer adaptador de entrada en C#.
- ⬜ **Modificar y guardar Word**
- ⬜ **Crear un archivo Word**

## Estructura de proyecto (actual)

```
shonoffice/
├── ShonOffice.sln
├── src/
│   ├── ShonOffice.Domain/          # C# — modelo de documento + puertos (interfaces). Sin dependencias externas.
│   │   ├── Documents/               #   OfficeDocument, WordDocument, Paragraph, TextRun, ParagraphAlignment, ExcelDocument, PowerPointDocument, Sheet, Slide
│   │   ├── Ports/                   #   IDocxReader/Writer, IXlsxReader/Writer, IPptxReader/Writer, IPdfEngine
│   │   └── Exceptions/              #   UnsupportedFormatException
│   ├── ShonOffice.Application/     # C# — casos de uso, depende solo de Domain
│   │   └── UseCases/                #   OpenDocumentUseCase, SaveDocumentUseCase, ConvertPdfToWordUseCase
│   ├── ShonOffice.Infra.OpenXml/   # C# — implementa IDocxReader con Open XML SDK
│   │   ├── WordOpenXmlReader.cs    #   lee .docx resolviendo formato real (no solo texto)
│   │   ├── StyleResolver.cs        #   recorre la cadena de estilos de Word (docDefaults → basedOn → directo)
│   │   └── PendingAdapters.cs      #   placeholders de IXlsxReader/IPptxReader (todavía no implementados)
│   ├── ShonOffice.Infra.Pdf/       # C# — implementa IPdfEngine 100% .NET ("C# por defecto", ver más arriba)
│   │   └── PendingAdapters.cs      #   placeholder de IPdfEngine (todavía no implementado)
│   └── ShonOffice.UI/              # C# + Avalonia — primer adaptador de entrada en C# (código, sin .axaml)
│       ├── Program.cs, App.cs      #   arranque de la app Avalonia
│       └── MainWindow.cs           #   abre .docx y lo renderiza con formato real
├── tests/
│   ├── ShonOffice.Application.Tests/    # casos de uso con adaptadores falsos (in-memory)
│   └── ShonOffice.Infra.OpenXml.Tests/  # valida la resolución de estilos con un .docx construido en memoria
├── legacy/
│   └── rust-docx-poc/               # Rust — fase inicial (GUI egui + lectura de .docx), PoC histórico
│       ├── Cargo.toml               #   crate autocontenido, no forma parte de ShonOffice.sln
│       ├── src/lib.rs, src/main.rs
│       └── README.md                #   por qué está acá y no en src/
└── README.md
```

`Domain` y `Application` siguen sin ninguna dependencia externa (ni Open
XML SDK, ni Avalonia, ni Rust) — a propósito, la regla de la arquitectura
hexagonal sigue siendo que las flechas de dependencia apuntan hacia
adentro. `Infra.OpenXml`, `Infra.Pdf` y `UI` sí dependen de librerías
externas porque son, justamente, los adaptadores pensados para eso.
`legacy/rust-docx-poc/` no depende de nada de lo anterior ni viceversa:
es un crate de Cargo aislado, fuera del build de `ShonOffice.sln`.

### Decisiones de diseño (fase actual)

- **C#/.NET es la tecnología principal**, según el "Nuevo rumbo
  arquitectónico" descripto más arriba. Toda funcionalidad nueva se
  evalúa primero en C#.
- **`ShonOffice.Infra.Pdf` es la implementación por defecto de
  `IPdfEngine`**, todavía como placeholder (`NotImplementedException`).
  No se introduce Rust para PDF hasta no tener evidencia concreta de que
  hace falta.
- **`legacy/rust-docx-poc/` se conserva sin mantenimiento activo**, solo
  como referencia histórica de la fase inicial (Rust + egui/eframe:
  binario pequeño, arranque instantáneo, sin runtime ni webview de por
  medio; `docx-rs` con `default-features = false` para no procesar
  imágenes embebidas solo para leer texto).

## Próximos pasos sugeridos

1. Extender `Infra.OpenXml` a `.xlsx` y `.pptx` (`NotImplementedExcelReader`
   y `NotImplementedPowerPointReader` son placeholders a propósito para
   ese paso — hoy solo lanzan `NotImplementedException`).
2. Implementar `IPdfEngine` en `ShonOffice.Infra.Pdf` con una librería
   .NET (por ejemplo PdfPig, iText o QuestPDF) — sigue siendo "C# por
   defecto": no se evalúa Rust para esto todavía.
3. Solo si, con datos reales, `Infra.Pdf` demuestra ser un cuello de
   botella (rendimiento, memoria, o una capacidad que el ecosistema .NET
   no cubre bien) se evalúa introducir `ShonOffice.Infra.Native` como
   una segunda implementación de `IPdfEngine` vía FFI a un motor Rust,
   siguiendo el criterio de "Nuevo rumbo arquitectónico". Hacerlo recién
   ahí evita que Rust se transforme en la pieza que define cómo se
   organiza todo lo demás.

