**English** · [Español](README.es.md)

# ooxmlkit

A low-level layer over `python-docx` that injects the OOXML the library does not cover: multilevel numbering bound to styles, hyphenation, Word fields, paragraph borders, and tables with no vertical rules. What it takes for a code-generated `.docx` to typeset like an editorial document instead of program output.

> **Closing the document requires Windows and Microsoft Word.** The module produces a valid, editable `.docx` on any system, but leaves the table of contents empty and the page total unresolved: those are fields only Word computes when it opens the file. The `cerrar.ps1` script updates them, embeds the fonts, and exports the PDF. Without that step the document works; its table of contents does not.

> The API is in Spanish. The functions are named `cuadro` (table), `figura` (figure), `parrafo` (paragraph), `sin_particion` (no hyphenation); this README documents each one.

## 1. What it is and who it is for

`python-docx` writes paragraphs, tables, and images. What it does not write is nearly everything that separates a typeset document from a typed one: section numbering that shows up in the table of contents and renumbers itself, hyphenation so justified text does not open rivers, a table of contents Word knows how to fill, tables with a thick rule on top and no vertical lines at all, suppressed hyphenation inside a three-centimetre column.

All of it exists in the OOXML format, and you reach it by injecting XML by hand into the tree `python-docx` maintains. `ooxmlkit` is that work already done: 647 lines, 34 functions, one external dependency.

It is for anyone generating Word documents from code that have to print well: recurring reports, technical documentation, normative deliverables. If the destination is a screen rather than paper, HTML solves it better and asks for less.

**If what you want is your own brand identity manual in Word,** start from the worked example rather than from the API: [`ejemplo/manual.py`](ejemplo/manual.py) typesets a complete seven-page manual for one person, and swapping her decisions for yours is the shortest path. The companion repository [identidad-personal](https://github.com/Carlos-Padilla-Bravo/identidad-personal) is what takes those decisions with you in the first place, and delivers the same manual in HTML and PDF; this one adds the editable `.docx`.

## 2. What it does / What it does not do

**What it does**

- **Real section numbering.** A multilevel list bound to the `Heading 1` and `Heading 2` styles, not numbers typed into the text. It appears in the table of contents, supports cross-references, and renumbers itself when a section is inserted. Appendices run in a separate lettered sequence.
- **Table of contents and Word fields.** Inserts `TOC`, `PAGE`, and `NUMPAGES` as fields, and flags the document so Word updates them on open.
- **Controlled hyphenation.** Turns on automatic hyphenation, which justified text needs, and suppresses it per style in the short measures where it looks bad: notes, cells, figure captions.
- **Typeset tables.** Numbered, caption above and note below, thick opening and closing rules, not one vertical line, fixed column widths, a header row that repeats when the table breaks across pages, and a last row that travels with its note.
- **Numbered figures** with caption and note, inserted at a width declared in centimetres.
- **WCAG 2.1 contrast, computed.** `ratio()` and `nivel()` return the contrast ratio and the level it reaches, so a document can state its contrast figures by calculation rather than by eye.
- **The palette as a dictionary of roles.** The ten keys are roles, not colours: `primario`, `senal`, `texto_sec`. Every function that takes a colour accepts a palette key or a bare hex value, so you can use the dictionary, replace it with `usar_paleta()`, or ignore it entirely.

**What it does not do**

- **It does not compute the table of contents or the pagination.** It leaves them as fields. Word resolves them, and that requires Windows. Said up front because it is the limitation that weighs most.
- **It is not a PyPI package.** It is a single-file module. You install it by copying it into your project.
- **It does not draw charts.** It inserts finished PNGs. The example generates its own with matplotlib, which the library neither uses nor requires.
- **It does not abstract `python-docx` away.** It complements it. You still build the document with `Document()`, `add_paragraph()`, and `add_table()`; `ooxmlkit` steps in where that API stops.
- **It does not handle templates, mail merge, or reading someone else's `.docx`.** Writing only.
- **It makes no API stability promise.** It is the code that typesets a real document, extracted so it can serve others. See the status section at the end.
- **It does not ship a design.** The default palette is a greyscale scaffold, there so the module typesets with no configuration. Choosing hues, type, and voice are identity decisions, and they are not a library's business: the companion repository settles those with the person.

## 3. Requirements

- **Python 3.9 or later** and **`python-docx`**: `pip install python-docx`. That is the only dependency. Tested with `python-docx` 1.2 on Python 3.14.
- **The fonts installed** on the machine that generates the document. The defaults are IBM Plex, under the OFL. If they are missing, Word substitutes silently. To use others, reassign them: `ooxmlkit.SERIF = "Georgia"`. The example reassigns its own three, Fraunces, Nunito Sans, and Space Mono, all OFL and available from Google Fonts; without them the manual still typesets, in another face.
- **Windows with Microsoft Word**, for `cerrar.ps1` only. Without it you still get the `.docx`, with an empty table of contents.

## 4. Installation

Copy `ooxmlkit.py` into your project, next to the code that will import it. There is no package and no installer: it is one file.

```bash
curl -O https://raw.githubusercontent.com/Carlos-Padilla-Bravo/ooxmlkit/main/ooxmlkit.py
```

Or clone the repository with GitHub Desktop and copy the file out.

## 5. Minimal use

```python
from docx import Document
from ooxmlkit import (construir_estilos, cuadro, numeracion_de_secciones,
                      pagina, parrafo, particion_de_palabras)

doc = Document()
construir_estilos(doc)          # 18 styles, already defined
particion_de_palabras(doc)      # required by justified text
numeracion_de_secciones(doc)    # numbers Heading 1 and Heading 2
pagina(doc.sections[0])         # A4, 3 cm and 2.5 cm margins

doc.add_heading("Resultados", level=1)
parrafo(doc, "Body text runs justified and hyphenated, with no rivers.")

cuadro(doc, 1, "Output by year",
       ["Year", "Units"],
       [["2024", "49"], ["2025", "63"]],
       anchos=[7.5, 7.5],
       nota="No vertical rules, thick rule above and below.")

doc.save("informe.docx")
```

To close it:

```powershell
powershell -File cerrar.ps1 informe.docx
```

## 6. The API

Three layers, most general first.

**OOXML layer.** What `python-docx` cannot write. It is the reason the module exists, and it knows nothing about the particular document it was written for.

| Function | What it does |
| --- | --- |
| `el(tag, **attrs)` | Builds an OOXML element with its attributes in Word's namespace |
| `campo(par, instruccion, texto_reserva)` | Inserts a field (`TOC`, `PAGE`, `NUMPAGES`) and returns the placeholder run |
| `actualizar_campos(doc)` | Flags the document so Word updates its fields on open |
| `particion_de_palabras(doc, zona_cm)` | Turns on automatic hyphenation, sparing capitals and capping consecutive hyphens at two |
| `sin_particion(pPr)` | Suppresses hyphenation in a `pPr`, for short measures |
| `idioma(doc, cod)` | Sets the document language, which governs hyphenation rules |
| `numeracion_de_secciones(doc)` | Multilevel list bound to `Heading 1` and `Heading 2`, plus a lettered appendix sequence. Returns the appendix id |
| `marcar_anexo(par, id_anexo)` | Moves a heading into the appendix sequence |
| `sin_numerar(par)` | Takes a heading out of all numbering |
| `pagina(seccion, izq, der, sup, inf)` | A4 and margins in centimetres |
| `portada_sin_encabezado(seccion)` | First page with no header or footer |
| `borde_par(par, lado, color, sz, espacio)` | Rule on one side of a paragraph |
| `sombra_par(par, color)` | Solid paragraph background |
| `sin_separar(par)` | Keeps the paragraph with the next, no widows or orphans |
| `salto(doc)` | Explicit page break |

**Content layer.** Typesetting for the elements that recur through a long document.

| Function | What it does |
| --- | --- |
| `texto(par, s, fuente, tam, color, negrita, cursiva)` | A run with its own formatting inside a paragraph |
| `parrafo(doc, s, estilo)` | Styled paragraph |
| `rico(doc, partes, estilo)` | Mixed-format paragraph: a list of strings or `(text, format)` pairs |
| `cuadro(doc, n, titulo, encabezados, filas, anchos, nota, alinear, fuente_datos)` | A complete numbered table. A cell starting with `**` goes bold; line breaks inside a cell are kept |
| `figura(doc, n, archivo, titulo, nota, ancho)` | Numbered figure with caption and note |
| `muestra_color(doc, claves, alto_cm)` | Solid colour band, typeset as a table rather than an image |

**Colour layer.** The palette and the contrast maths.

| Function | What it does |
| --- | --- |
| `PALETA` | Ten roles: `texto`, `texto_fuerte`, `texto_sec`, `texto_ter`, `primario`, `primario_int`, `senal`, `fondo`, `superficie`, `borde`. The defaults are a greyscale scaffold, not a design |
| `usar_paleta(colores)` | Adds or replaces colours. Only touches the keys you pass |
| `hexa(k)` | Hex without the hash, from a key or a bare hex value. Raises `ValueError` on an invalid colour, so a mistyped key cannot slip into the XML silently |
| `hx(k)` | Uppercase hex with the hash, for printing inside the document |
| `rgb(k)` | The colour as a `python-docx` `RGBColor` |
| `rgb_txt(k)` · `hsl_txt(k)` | The colour as text, in RGB or HSL |
| `lum(k)` | WCAG 2.1 relative luminance |
| `ratio(a, b)` | Contrast ratio between two colours, 1 to 21 |
| `nivel(v, grande)` | The level that ratio reaches: `AAA`, `AA`, `AA grande`, or `no cumple` |
| `contra(a, b)` | The ratio as text, with a decimal comma |

**Worked example.** `construir_estilos(doc)` defines 18 styles with their families, sizes, leading, and colours already decided, and restyles the four heading levels and `Normal` on top of that. They are `Cuerpo`, `Entradilla`, `Destacado`, `Nota`, `Dato`, `TituloCuadro`, `PieFigura`, `NotaTabla`, `Vineta`, `Enum`, `Regla`, `Imagen`, `Cita`, `Encabezado`, and four cover styles. It is not a configurable API: it is the closed typographic scheme of the document the module was written for, offered as a starting point. The same goes for `claro_oscuro()` and `nom_hx()`, which build rows of one particular spec table. For another document you copy those functions and edit them.

## 7. The example

`ejemplo/manual.py` typesets **Ignacia Fuentes's brand identity manual**: seven pages with a cover, table of contents, three numbered sections, a lettered appendix, four tables, a figure, a colour band, and a footer with page numbering. Ignacia is a fictional person, and her system is published in the companion repository [identidad-personal](https://github.com/Carlos-Padilla-Bravo/identidad-personal), where the same manual comes out as HTML and PDF. Same person in both places: here you see what the `.docx` adds.

The example is split across two files on purpose, and that split is the part worth copying:

- **`ejemplo/ignacia.py`** — what the person decides: her ten colour roles, her three families, two notes. No OOXML.
- **`ejemplo/manual.py`** — what the library typesets. Nothing about Ignacia beyond what it imports from the file above.

Replacing the first with your own and adjusting the prose in the second is the short path to your own manual. The two lines that turn the library into one person's system are these, and they go before `construir_estilos()`:

```python
ooxmlkit.usar_paleta(PALETA)                  # the ten roles, with your values
ooxmlkit.SERIF, ooxmlkit.SANS = SERIF, SANS   # your families
```

The contrast figures in Table 2 are computed at generation time with `ratio()`, so no value is typed by hand: change a colour and the table follows.

```bash
python ejemplo/figura.py           # the PNG, only if you want to regenerate it
python ejemplo/manual.py           # the .docx, from any directory
powershell -File cerrar.ps1 ejemplo\manual.docx
```

The output is in the repository: [`manual.docx`](ejemplo/manual.docx) · [`manual.pdf`](ejemplo/manual.pdf).

## 8. Known traps

Things that cost an afternoon if nobody tells you.

- **Order matters when inserting into `w:pPr`.** The OOXML schema fixes the sequence of children. An element appended at the end saves without error and Word ignores it silently. That is why `sin_particion()` and the numbering use `insert_element_before()` with the successor list rather than `append()`.
- **Page breaks belong in the style, not in an empty paragraph.** `Heading 1` carries `page_break_before`. An empty paragraph holding a break can land at the foot of a page and leave a blank one behind.
- **If a font is not installed, Word substitutes silently.** The document generates anyway and typesets in another face. Worth checking in the PDF.
- **Table columns are fixed width.** A value that does not fit gets cut mid-word. When you change a table's content, check the PDF.
- **Typesetting defects do not show up in extracted text.** Overflows, orphans, narrow columns, and ugly hyphenation are only visible by looking at the PDF page by page.

## 9. Licence and status

Published under the **MIT** licence. Copyright (c) 2026 Carlos Padilla Bravo. You may use, copy, and modify it, including in paid work, keeping the attribution notice.

**Status: occasional maintenance.** The module was extracted from a project in use and gets updated when that project demands it. There is no promise of support or response times, and issues are disabled, so this is not a support channel. Fork it without asking: that is what the MIT licence above is for.

---

Author: **Carlos Padilla Bravo**

`ooxmlkit` came out of the toolchain that generates a personal brand identity manual. The two repositories divide the work like this:

| | [identidad-personal](https://github.com/Carlos-Padilla-Bravo/identidad-personal) | ooxmlkit |
| --- | --- | --- |
| What it is | A Claude Code skill | A Python library |
| What it settles | **What the person decides**: essence, territory, colour, type, voice | **How the document is typeset** |
| What it delivers | The manual in HTML, and its PDF from the browser | The manual as an editable `.docx` |
| What it asks for | Nothing beyond a browser | Windows with Word, to close it |

The natural order is one then the other: the skill takes the decisions with you, and this library typesets them in Word. Neither uses the other, which is why the example here and the example there are the same person.
