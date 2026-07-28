[English](README.md) · **Español**

# ooxmlkit

Una capa de bajo nivel sobre `python-docx` que inyecta el OOXML que la librería no cubre: numeración multinivel ligada a estilos, partición de palabras, campos de Word, bordes de párrafo y cuadros sin líneas verticales. Lo que hace falta para que un `.docx` generado por código se componga como un documento editorial y no como una salida de programa.

> **Requiere Windows y Microsoft Word para cerrar el documento.** El módulo genera un `.docx` válido y editable en cualquier sistema, pero deja el índice vacío y sin total de páginas: son campos que solo Word resuelve al abrir el archivo. El script `cerrar.ps1` los actualiza, incrusta las tipografías y exporta el PDF. Sin ese paso, el documento sirve; su índice no.

> La API está en español. Los nombres de las funciones son `cuadro`, `figura`, `parrafo`, `sin_particion`; este README los documenta uno por uno.

## 1. Qué es y a quién sirve

`python-docx` escribe párrafos, tablas e imágenes. Lo que no escribe es casi todo lo que distingue un documento compuesto de uno tecleado: una numeración de secciones que aparezca en el índice y se renumere sola, partición de palabras para que el texto justificado no abra ríos, un índice que Word sepa rellenar, cuadros con filete grueso arriba y ninguna línea vertical, supresión de partición dentro de una columna de tres centímetros.

Todo eso existe en el formato OOXML, y se llega a ello inyectando XML a mano en el árbol que `python-docx` mantiene. `ooxmlkit` es ese trabajo ya hecho: 632 líneas, 34 funciones, una sola dependencia externa.

Sirve a quien genera documentos Word por código y necesita que se impriman bien: informes recurrentes, documentación técnica, entregables normativos. Si el destino es la pantalla y no el papel, HTML resuelve mejor y con menos requisitos.

## 2. Qué hace / Qué no hace

**Qué hace**

- **Numeración de secciones de verdad.** Una lista multinivel ligada a los estilos `Heading 1` y `Heading 2`, no números escritos en el texto. Aparece en el índice, admite referencias cruzadas y se renumera sola al insertar una sección. Los anexos van en una secuencia aparte, con letra.
- **Índice y campos de Word.** Inserta `TOC`, `PAGE` y `NUMPAGES` como campos, y marca el documento para que Word los actualice al abrirlo.
- **Partición de palabras controlada.** Activa la partición automática, que el texto justificado exige, y la suprime por estilo en las medidas cortas donde afea: notas, celdas, pies de figura.
- **Cuadros compuestos.** Numerados, con título arriba y nota al pie, filete grueso de apertura y cierre, sin una sola línea vertical, ancho de columna fijo y encabezado que se repite al partir el cuadro entre páginas.
- **Figuras numeradas** con pie y nota, insertadas a un ancho declarado en centímetros.
- **Contraste WCAG 2.1 calculado.** `ratio()` y `nivel()` devuelven la razón de contraste y el nivel que alcanza, para que un documento pueda declarar sus contrastes por cálculo y no a ojo.
- **Paleta como diccionario.** Cualquier función que recibe un color acepta una clave de la paleta o un hexadecimal suelto, así que se puede usar el diccionario, sustituirlo con `usar_paleta()` o ignorarlo por completo.

**Qué no hace**

- **No calcula el índice ni la paginación.** Los deja como campos. Los resuelve Word, y eso obliga a Windows. Está dicho arriba porque es la limitación que más pesa.
- **No es un paquete de PyPI.** Es un módulo de un archivo. Se instala copiándolo al proyecto.
- **No dibuja gráficos.** Inserta PNG ya compuestos. El ejemplo genera el suyo con matplotlib, que la librería no usa ni requiere.
- **No abstrae `python-docx`.** Lo complementa. El documento se sigue armando con `Document()`, `add_paragraph()` y `add_table()`; `ooxmlkit` interviene donde esa API se queda corta.
- **No maneja plantillas, combinación de correspondencia ni lectura de `.docx` ajenos.** Solo escritura.
- **No promete estabilidad de API.** Es el código que compone un documento real, extraído para que sirva a otros. Ver el estado, al final.

## 3. Requisitos

- **Python 3.9 o superior** y **`python-docx`**: `pip install python-docx`. Es la única dependencia. Probado con `python-docx` 1.2 y Python 3.14.
- **Las tipografías instaladas** en el equipo que genera el documento. Las familias por defecto son IBM Plex, de licencia OFL. Si no están, Word sustituye sin avisar. Para usar otras, se reasignan: `ooxmlkit.SERIF = "Georgia"`.
- **Windows con Microsoft Word**, solo para `cerrar.ps1`. Sin esto se obtiene el `.docx`, con el índice vacío.

## 4. Instalación

Copiar `ooxmlkit.py` al proyecto, junto al código que lo va a importar. No hay paquete ni instalador: es un archivo.

```bash
curl -O https://raw.githubusercontent.com/Carlos-Padilla-Bravo/ooxmlkit/main/ooxmlkit.py
```

O clonar el repositorio con GitHub Desktop y copiar el archivo.

## 5. Uso mínimo

```python
from docx import Document
from ooxmlkit import (construir_estilos, cuadro, numeracion_de_secciones,
                      pagina, parrafo, particion_de_palabras)

doc = Document()
construir_estilos(doc)          # 18 estilos ya definidos
particion_de_palabras(doc)      # exigida por el texto justificado
numeracion_de_secciones(doc)    # numera Heading 1 y Heading 2
pagina(doc.sections[0])         # A4 con márgenes de 3 y 2,5 cm

doc.add_heading("Resultados", level=1)
parrafo(doc, "El cuerpo va justificado, con partición y sin ríos de espacio.")

cuadro(doc, 1, "Producción por año",
       ["Año", "Unidades"],
       [["2024", "49"], ["2025", "63"]],
       anchos=[7.5, 7.5],
       nota="Sin líneas verticales, con filete de apertura y cierre.")

doc.save("informe.docx")
```

Para cerrarlo:

```powershell
powershell -File cerrar.ps1 informe.docx
```

## 6. La API

Tres capas, de más a menos general.

**Capa OOXML.** Lo que no se puede escribir desde `python-docx`. Es el motivo de que exista el módulo, y no sabe nada del documento concreto para el que se escribió.

| Función | Qué hace |
| --- | --- |
| `el(tag, **attrs)` | Crea un elemento OOXML con sus atributos en el espacio de nombres de Word |
| `campo(par, instruccion, texto_reserva)` | Inserta un campo (`TOC`, `PAGE`, `NUMPAGES`) y devuelve el run del texto de reserva |
| `actualizar_campos(doc)` | Marca el documento para que Word actualice los campos al abrirlo |
| `particion_de_palabras(doc, zona_cm)` | Activa la partición automática, sin partir mayúsculas ni más de dos líneas seguidas |
| `sin_particion(pPr)` | Suprime la partición en un `pPr`, para medidas cortas |
| `idioma(doc, cod)` | Fija el idioma del documento, que decide las reglas de partición |
| `numeracion_de_secciones(doc)` | Lista multinivel ligada a `Heading 1` y `Heading 2`, más una secuencia de anexos con letra. Devuelve el identificador de anexos |
| `marcar_anexo(par, id_anexo)` | Pasa un título a la secuencia de anexos |
| `sin_numerar(par)` | Saca un título de toda numeración |
| `pagina(seccion, izq, der, sup, inf)` | A4 y márgenes en centímetros |
| `portada_sin_encabezado(seccion)` | Primera página sin encabezado ni pie |
| `borde_par(par, lado, color, sz, espacio)` | Filete en un lado del párrafo |
| `sombra_par(par, color)` | Fondo sólido de párrafo |
| `sin_separar(par)` | Mantiene el párrafo con el siguiente, sin viudas ni huérfanas |
| `salto(doc)` | Salto de página explícito |

**Capa de contenido.** Composición de los elementos que se repiten en un documento largo.

| Función | Qué hace |
| --- | --- |
| `texto(par, s, fuente, tam, color, negrita, cursiva)` | Un run con formato propio dentro de un párrafo |
| `parrafo(doc, s, estilo)` | Párrafo con estilo |
| `rico(doc, partes, estilo)` | Párrafo mezclando formatos: lista de cadenas o de `(texto, formato)` |
| `cuadro(doc, n, titulo, encabezados, filas, anchos, nota, alinear, fuente_datos)` | Cuadro numerado completo. Una celda que empieza con `**` va en negrita; los saltos de línea dentro de una celda se respetan |
| `figura(doc, n, archivo, titulo, nota, ancho)` | Figura numerada con pie y nota |
| `muestra_color(doc, claves, alto_cm)` | Franja de color sólida, compuesta como tabla y no como imagen |

**Capa de color.** La paleta y el cálculo de contraste.

| Función | Qué hace |
| --- | --- |
| `PALETA` | Diccionario de diez colores por defecto |
| `usar_paleta(colores)` | Agrega o sustituye colores. Solo toca las claves que se pasan |
| `hexa(k)` | Hexadecimal sin almohadilla, de una clave o de un hex suelto. Levanta `ValueError` si el color no es válido, para que una clave mal escrita no entre al XML en silencio |
| `hx(k)` | Hexadecimal en mayúsculas y con almohadilla, para imprimirlo en el documento |
| `rgb(k)` | El color como `RGBColor` de `python-docx` |
| `rgb_txt(k)` · `hsl_txt(k)` | El color como texto, en RGB o en HSL |
| `lum(k)` | Luminancia relativa WCAG 2.1 |
| `ratio(a, b)` | Razón de contraste entre dos colores, de 1 a 21 |
| `nivel(v, grande)` | Nivel que alcanza esa razón: `AAA`, `AA`, `AA grande` o `no cumple` |
| `contra(a, b)` | La razón como texto, con coma decimal |

**Ejemplo trabajado.** `construir_estilos(doc)` define 18 estilos con sus familias, cuerpos, interlineados y colores ya decididos, y reescribe además los cuatro niveles de título y `Normal`. Son `Cuerpo`, `Entradilla`, `Destacado`, `Nota`, `Dato`, `TituloCuadro`, `PieFigura`, `NotaTabla`, `Vineta`, `Enum`, `Regla`, `Imagen`, `Cita`, `Encabezado` y los cuatro de portada. No es una API configurable: es la composición cerrada del documento para el que se escribió el módulo, y sirve como punto de partida. Lo mismo vale para `claro_oscuro()` y `nom_hx()`, que arman filas de una ficha técnica concreta. Para otro documento, esas funciones se copian y se editan.

## 7. El ejemplo

`ejemplo/informe.py` genera un informe de cinco páginas que ejercita toda la API: portada, índice, dos secciones numeradas, un anexo con letra, dos cuadros, una figura, una franja de color y un pie con numeración de páginas. El cuadro del anexo se compone leyendo la paleta y calculando los contrastes al generar, así que ningún valor está escrito a mano.

```bash
python ejemplo/figura.py           # el PNG, solo si se quiere regenerar
python ejemplo/informe.py          # el .docx, desde la raíz del repositorio
powershell -File cerrar.ps1 ejemplo\informe.docx
```

El resultado está en el repositorio: [`informe.docx`](ejemplo/informe.docx) · [`informe.pdf`](ejemplo/informe.pdf).

## 8. Trampas conocidas

Cosas que cuestan una tarde si no están dichas.

- **Al insertar en `w:pPr`, el orden importa.** El esquema de OOXML fija la secuencia de los hijos. Un elemento agregado al final se guarda sin error y Word lo ignora en silencio. Por eso `sin_particion()` y la numeración usan `insert_element_before()` con la lista de sucesores, no `append()`.
- **Los saltos de página van en el estilo, no en un párrafo vacío.** `Heading 1` lleva `page_break_before`. Un párrafo vacío con salto puede caer al pie de la página y dejar una página en blanco.
- **Si la tipografía no está instalada, Word sustituye sin avisar.** El documento se genera igual y se compone con otra letra. Conviene comprobarlo en el PDF.
- **Las columnas de los cuadros son de ancho fijo.** Un valor que no cabe se corta a mitad de palabra. Al cambiar el contenido de un cuadro, revisar el PDF.
- **Los defectos de composición no aparecen en el texto extraído.** Desbordes, huérfanos, columnas estrechas y particiones feas solo se ven mirando el PDF página por página.

## 9. Licencia y estado

Publicado bajo licencia **MIT**. Copyright (c) 2026 Carlos Padilla Bravo. Se puede usar, copiar y modificar, incluso en trabajo pagado, conservando el aviso de autoría.

**Estado: mantención ocasional.** El módulo se extrajo de un proyecto en uso y se actualiza cuando ese proyecto lo exige. No hay promesa de soporte ni de tiempos de respuesta, y los issues están desactivados, así que este no es un canal de soporte. Se puede hacer un fork sin pedir permiso: para eso está la licencia MIT.

---

Autor: **Carlos Padilla Bravo**

`ooxmlkit` salió de la cadena que genera un manual de identidad de marca personal. El método de ese manual está publicado aparte, como skill de Claude Code: [identidad-personal](https://github.com/Carlos-Padilla-Bravo/identidad-personal). Esa skill entrega HTML y su PDF, y no usa esta librería; `ooxmlkit` importa cuando el entregable tiene que ser un `.docx` editable.
