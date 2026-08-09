# ShonOffice

Suite de ofimática liviana para leer y (más adelante) editar archivos de
Microsoft Office (`.xlsx`, `.docx`, `.pptx`), con conversión PDF↔Word.

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

El repo hoy tiene **tres partes** que conviven mientras se migra de una a
otra (ver "Estado actual" más abajo). No hay un único comando que las
corra a todas: se ejecutan por separado.

### Parte 1 (recomendada) — UI en C# (Avalonia + Open XML SDK) — lee `.docx` con formato real

Esta es la UI que hay que usar hoy para ver un `.docx` como se ve en
Word de verdad: usa `ShonOffice.Infra.OpenXml` (Open XML SDK) para leer
el documento con su formato real y `ShonOffice.UI` (Avalonia) para
mostrarlo — ver "Por qué la GUI no se veía como Word real" más arriba.

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


## Arquitectura de lenguajes: C# + Rust vía FFI

ShonOffice usa una **arquitectura híbrida C# + Rust**, donde cada lenguaje
resuelve lo que mejor sabe hacer:

- **C# (.NET) + Avalonia** — UI y toda la lectura/escritura de `.xlsx`,
  `.docx` y `.pptx` mediante **Open XML SDK**, la librería oficial de
  Microsoft para OOXML. Se eligió C# para esta capa porque Open XML SDK es
  muchísimo más completo y maduro que cualquier alternativa en Rust, en
  particular para PowerPoint, donde el ecosistema Rust no tiene una
  librería a la altura.

- **Rust**, compilado como librería nativa (`cdylib`) y expuesto vía
  **FFI** (`P/Invoke` / `DllImport` desde C#, generando los bindings con
  [`csbindgen`](https://github.com/Cysharp/csbindgen) en vez de
  escribirlos a mano) — se encarga de la carga pesada, principalmente:
  - Parsing y reconstrucción de PDF (extracción de texto/layout para la
    conversión PDF→Word).
  - Cualquier otro procesamiento intensivo que se identifique más
    adelante.

**Por qué FFI y no otra cosa:** se descartó Rust puro porque no existe un
equivalente serio a Open XML SDK para `.pptx`. Se descartó C# puro porque
el parsing/reconstrucción de PDF es trabajo pesado donde Rust rinde mejor
y da más control de memoria. FFI directo (linkeo de la `cdylib` dentro del
proceso de C#) es la opción por defecto por menor overhead; si el
acoplamiento resulta problemático (crashes difíciles de depurar,
marshaling), la alternativa es Rust como **proceso separado** comunicado
por IPC (stdin/stdout con JSON, o named pipes), con más aislamiento a
costa de algo de rendimiento.

## Arquitectura de software para escalar: Hexagonal (Ports & Adapters)

Elegir C#+Rust resuelve *qué lenguaje hace qué*, pero no resuelve *cómo se
organiza el código para poder crecer* (sumar xlsx, pptx, edición,
conversión PDF, y eventualmente otra UI o un modo CLI) sin que todo
termine acoplado a Open XML SDK, a Avalonia, o al binario de Rust. Para
eso propongo **arquitectura hexagonal (Ports & Adapters)**, con el
dominio en el centro y todo lo externo (librerías, UI, FFI) como piezas
reemplazables alrededor.

### Por qué hexagonal y no otra cosa

- **Los formatos Office cambian de librería más seguido que la lógica de
  negocio.** Si hoy Open XML SDK resuelve todo pero mañana hace falta otra
  librería para un caso puntual (o un formato viejo, `.doc`/`.xls`), el
  dominio (modelo de documento, casos de uso) no debería enterarse. Eso
  solo es alcanzable si el dominio depende de **interfaces propias**
  (puertos) y no directamente de Open XML SDK.
- **El motor Rust es, ni más ni menos, otro adaptador.** Desde el punto de
  vista del dominio, "convertir PDF a Word" es un caso de uso que depende
  de un puerto `IPdfEngine`; que ese puerto hoy lo implemente FFI a Rust
  (y mañana, si hiciera falta, un proceso separado por IPC) es un detalle
  de infraestructura que no debería filtrarse a la lógica de negocio ni a
  la UI.
- **Multiplica los puntos de entrada sin duplicar lógica.** Hoy hay GUI y
  CLI; el dominio y los casos de uso ya son compartidos vía `lib.rs` en la
  fase Rust. Hexagonal formaliza ese mismo principio para la fase C#: UI,
  CLI, o incluso un futuro servicio/API son todos "adaptadores de
  entrada" que llaman a los mismos casos de uso.
- **Testeable sin abrir archivos reales.** Los casos de uso se pueden
  probar con adaptadores falsos (in-memory) para xlsx/docx/pptx y para el
  motor Rust, sin depender de Open XML SDK ni de compilar la `cdylib` en
  cada corrida de tests.
- **Costo de entrada bajo para un proyecto de este tamaño.** Se descartó
  microservicios (no hay necesidad de escalar por separado ni de
  despliegue distribuido) y una arquitectura en capas estrictamente
  lineal (UI→Lógica→Datos) porque ahí Open XML SDK y el FFI a Rust
  terminan filtrándose hacia arriba tarde o temprano. Hexagonal da el
  aislamiento necesario con una curva de aprendizaje razonable.

### Capas y flujo de dependencias

```
                     ┌─────────────────────────────┐
   Adaptadores de    │   ShonOffice.UI (Avalonia)   │
   entrada           │   ShonOffice.Cli             │
                     └──────────────┬───────────────┘
                                    │ llama a
                     ┌──────────────▼───────────────┐
   Aplicación        │   ShonOffice.Application      │
   (casos de uso)    │   - AbrirDocumento            │
                     │   - ConvertirPdfAWord         │
                     │   - GuardarDocumento           │
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
              ┌─────────────────────┼─────────────────────┐
              │                     │                      │
┌─────────────▼───────────┐ ┌───────▼────────────┐ ┌───────▼──────────┐
│ ShonOffice.Infra.OpenXml │ │ ShonOffice.Infra.Native  │ │ (futuros adaptadores) │
│ - implementa IDocxReader │ │ - bindings csbindgen     │ │ - ej. exportar a ODF  │
│   con Open Xml SDK       │ │ - implementa IPdfEngine  │ │   sin tocar el resto  │
└──────────────────────────┘ │   llamando a la cdylib   │ └───────────────────────┘
                              └───────────┬───────────────┘
                                          │ FFI
                              ┌───────────▼───────────────┐
                              │ shonoffice-native (Rust)   │
                              │ - cdylib                   │
                              │ - parsing/reconstrucción    │
                              │   de PDF                    │
                              └────────────────────────────┘
```

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
│   ├── ShonOffice.Infra.Native/    # C# — implementa IPdfEngine llamando a la cdylib (bindings csbindgen)
│   ├── ShonOffice.UI/              # C# + Avalonia — adaptador de entrada (MVVM)
│   ├── ShonOffice.Cli/             # C# — adaptador de entrada (terminal)
│   └── shonoffice-native/          # Rust — cdylib, parsing/reconstrucción de PDF
│       ├── src/
│       └── bindings/               # bindings generados con csbindgen
├── tests/
│   ├── ShonOffice.Application.Tests/  # casos de uso con adaptadores falsos (in-memory)
│   └── shonoffice-native.tests/       # tests del motor Rust
└── README.md
```

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
│   ├── ShonOffice.UI/              # C# + Avalonia — primer adaptador de entrada en C# (código, sin .axaml)
│   │   ├── Program.cs, App.cs      #   arranque de la app Avalonia
│   │   └── MainWindow.cs           #   abre .docx y lo renderiza con formato real
│   ├── lib.rs, main.rs, bin/cli.rs # Rust (fase inicial) — GUI + CLI que ya leen .docx, quedan como referencia/PoC
│   └── (Infra.Native, Cli, shonoffice-native: aún no creados)
├── tests/
│   ├── ShonOffice.Application.Tests/    # casos de uso con adaptadores falsos (in-memory)
│   └── ShonOffice.Infra.OpenXml.Tests/  # valida la resolución de estilos con un .docx construido en memoria
├── Cargo.toml                      # sigue compilando la fase Rust independientemente
└── README.md
```

`Domain` y `Application` siguen sin ninguna dependencia externa (ni Open
XML SDK, ni Avalonia, ni el motor Rust) — a propósito, la regla de la
arquitectura hexagonal sigue siendo que las flechas de dependencia
apuntan hacia adentro. `Infra.OpenXml` y `UI` sí dependen de librerías
externas porque son, justamente, los adaptadores pensados para eso.

```

### Decisiones de diseño (fase actual)

- **Rust + egui/eframe** para esta fase: binario pequeño, arranque
  instantáneo, sin runtime ni webview de por medio.
- `docx-rs` con `default-features = false`: no se necesita procesar
  imágenes embebidas solo para leer texto.
- `profile.release` ajustado (`opt-level = "z"`, `lto = true`, `strip =
  true`, `panic = "abort"`) para un binario final lo más pequeño posible.

## Próximos pasos sugeridos

2. Extender `Infra.OpenXml` a `.xlsx` y `.pptx` (`NotImplementedExcelReader`
   y `NotImplementedPowerPointReader` son placeholders a propósito para
   ese paso — hoy solo lanzan `NotImplementedException`).
3. Recién ahí encarar el FFI: exponer `shonoffice-native` como `cdylib`,
   generar bindings con `csbindgen`, e implementar `IPdfEngine` en
   `Infra.Native`. Hacerlo después de tener el esqueleto hexagonal evita
   que el FFI se transforme en la pieza que define cómo se organiza todo
   lo demás.
