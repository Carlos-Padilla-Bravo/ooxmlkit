# CLAUDE.md

Guidance for working **on** this repository in Claude Code. It orients whoever
clones the repo to modify it; it is not needed to merely use the module.

## What this repo is

One file, `ooxmlkit.py`, plus the example that exercises it. The module is a
low-level layer over `python-docx` that injects the OOXML `python-docx` does not
cover. Both READMEs document every public function; read `README.md` first and
don't restate it here.

## Decisions already closed

Don't reopen these without being asked.

- **The API stays in Spanish.** Renaming it would either force a rewrite of the
  call sites in the project this was extracted from, or leave the published copy
  diverging from the one actually in use. No English aliases either: two names
  for one thing is the duplication the module otherwise avoids.
- **Single file, no package.** No `setup.py`, no `pyproject.toml`, no PyPI. You
  install it by copying it. Adding packaging is a different project.
- **`python-docx` is the only dependency.** The WCAG contrast maths lives in the
  module for that reason. Don't import a colour library, and don't reintroduce a
  separate palette module: the palette is the `PALETA` dict and `usar_paleta()`.
- **`construir_estilos()`, `claro_oscuro()` and `nom_hx()` are a worked example,
  not an API.** They are the closed typographic scheme of one document. Don't
  turn them into something configurable; that is what copying and editing is for.
- **The Windows and Word requirement is declared on the first screen of both
  READMEs.** It is the limitation that weighs most. Keep it there; don't move it
  into a footnote.
- **LibreOffice is not offered as a substitute for `cerrar.ps1`.** It has never
  been tested here, so it can be mentioned as unverified or not at all.

## Working conventions

- **Order matters when inserting into `w:pPr`.** The OOXML schema fixes the
  sequence of children. An element appended at the end saves without error and
  Word ignores it silently. Use `insert_element_before()` with the successor
  tuple (`_TRAS_NUMPR`, `_TRAS_PARTICION`), never `append()`.
- **`hexa()` validates on purpose.** Without it a mistyped palette key travels
  into the XML as its own name and Word discards it without complaint: a wrong
  document with nothing having failed. Don't relax it. `"auto"` stays legal
  because OOXML accepts it as a colour value.
- **Page breaks belong in a style, not in an empty paragraph.** An empty
  paragraph carrying a break can land at the foot of a page and leave a blank one.
- **Keep both READMEs in step.** `README.md` (English, GitHub's default) and
  `README.es.md` carry the same content, the same section count and the same
  figures; edit them together.
- **The figures quoted in the READMEs are checked against the code**, not
  estimated: line count, number of public functions, number of styles, number of
  palette entries, pages in the example. If you change the module, recount them.
- **Regenerating the example rewrites a 4.8 MB blob.** `ejemplo/informe.docx`
  carries 16 embedded font files, which is why it opens correctly on a machine
  without IBM Plex. Regenerate it when the output really changes, not by reflex:
  every regeneration adds another copy to the history.
- **A copy of this module still runs inside the private project it came from.**
  A fix here usually needs porting there, and the other way round. They are two
  files, not one; assume they have drifted until checked.

## Layout

- `ooxmlkit.py` — the module.
- `cerrar.ps1` — drives Word over COM to resolve fields, embed fonts, export PDF.
- `ejemplo/informe.py` — the five-page report that exercises the whole API;
  `ejemplo/figura.py` regenerates its PNG and is the only thing here that wants
  matplotlib.
- `README.md` / `README.es.md`, `LICENSE`.

## Verifying a change

There is no test suite; the deliverable is checked on the output.

```bash
python ejemplo/informe.py
powershell -File cerrar.ps1 ejemplo\informe.docx
```

Then read the PDF page by page. Typesetting defects do not appear in extracted
text: overflows, orphans, narrow columns and ugly hyphenation are only visible by
looking. Check no blank pages, that the table of contents carries real page
numbers, and that section numbering and the appendix letter still render.

Exercise the hostile cases too, not only the happy path: a table with one row and
with none, text with accents and ñ, a mistyped colour key (which must raise), a
replaced palette, and the example run from a directory other than the repo root.

## Status

Maintained occasionally, with no promise of support. It changes when the project
it was extracted from demands it. Issues are disabled, so this is not a support
channel.
