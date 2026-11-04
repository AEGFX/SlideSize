# List Builder

Turns a spreadsheet of names into a PowerPoint deck of lists, N names a slide. The name boxes are real placeholders that inherit from a LIST layout on one slide master, so the whole deck is restyled from Slide Master view. Typical job is a dinner table plan, 300 guests, 30 a slide, sorted by surname. Everything runs in the browser. Nothing is uploaded.

Lives at `/list-builder.html` with the engine in `list-builder.js`. The homepage links to it under PPT Tools.

## How the pptx is written and why

The deck had to pass one test. Open it in PowerPoint, go to Slide Master, move or restyle a name box on the LIST layout, every slide follows. That only works when the layout holds the placeholders and each slide shape references the layout placeholder by `idx` with no position and no run formatting of its own.

PptxGenJS was tried first (3.12.0, `defineSlideMaster` with placeholder objects, `addSlide({masterName})` and `addText(text, {placeholder})`). It does write `<p:ph type="body" idx="N"/>` on both the layout and the slide, but every slide shape also gets its own `<a:xfrm>` with the position copied from the layout and its own `<a:rPr sz="1800">`, and the shape names come out as `Text 0`. Move the box on the layout and the slides stay where they were. That is not inheritance, so PptxGenJS was dropped.

The package is written by hand with JSZip instead, which SlideSize already uses elsewhere. The NAMEDROP reference deck was unzipped and used as the structural model for the package skeleton, master and manifest. The placeholders go further than the reference, which copies plain text boxes onto every slide. What this writer produces

- one master, white background, Calibri through the theme fonts, body style with no bullets
- a layout named `LIST`, and `LIST BREAK` as well when breaking on table, each holding every placeholder with its position, size, alignment, font size and prompt text
- one slide per chunk whose shapes are `<p:sp>` with `<p:ph type="body" idx="N"/>`, an empty `<p:spPr/>`, an empty `<a:bodyPr/>` and the text only. Nothing else. Position, size and formatting all come from the layout
- the title is a `<p:ph type="title"/>` so it takes the master title style
- every placeholder carries its NAMEDROP shape name, `row.1.name`, `row.1.first_name`, `row.1.last_name`, `row.1.table`, `row.1.extra`, `header`, `title`, `range.start`, `range.end`, `page.number`, `page.count`
- a hidden text box named `NAMEDROP.MANIFEST` on the master, off the left edge of the canvas, holding the same JSON scaffold.py writes plus a `generator` key, so the desktop app can recognise the deck later
- the static words `to` and `of` are plain text boxes on the layout

The structure was checked with python-pptx and by reading the XML. Every slide shape resolves to a layout placeholder with the same idx and name, and python-pptx reports the slide shape position as the layout's position because the slide has none of its own. LibreOffice renders it correctly. The one thing that still needs a human is the restyle test in PowerPoint itself, see Checking in PowerPoint below.

## Using it

1. Drop an .xlsx or .csv. The header row is detected and can be changed. A workbook with several sheets gets a sheet picker.
2. Check the three column dropdowns, First name, Last name, Table. A single Full name column can be mapped instead of first and last. Any other column can be shown as an extra field next to each name.
3. Set names per slide, the sort, breaking on table change, UK or Dutch particles, one name box or separate first and last boxes, an optional title and whether to show the range and the page count.
4. Look at the preview of the first two slides.
5. Press Build deck. The file is named from the source file plus a timestamp.

The warnings list sits under the preview. Nothing in it blocks the build.

## What the engine does

A port of the NAMEDROP Flow engine (engine_flow.py, rules.py, source.py) to JavaScript, the same code in the page and in the tests.

Reading the file. Cells are cleaned of leading and trailing whitespace, tabs, non-breaking spaces, zero width characters, byte order marks and line breaks, with double spaces collapsed. Each cleaning is reported as info. Hidden rows in a workbook are skipped. A formula with no saved result reads as blank and is reported. CSV is decoded as UTF-8 with or without BOM, falling back to Windows-1252, and the delimiter is sniffed from the first 4 KB (comma, semicolon or tab). Quoted fields with embedded newlines and doubled quotes are handled.

Header row. The first row where every non-empty cell is text, at least 60 percent of the widest row is filled, no value repeats within the row, no value repeats below it in its own column, and the next non-empty row has a different type signature. The candidates are kept so the user can override.

Column guesses. By header name first (first name, forename, given name, surname, last name, family name, table, table number, and the Dutch voornaam, achternaam, tafel), then by a header that contains one of those as whole words (Guest First Name, Delegate Surname), then for the table by content, a column that is more than 80 percent small integers.

Names. First and last from two columns, or one full name column split into first and surname where the surname is the last word plus any lower case particle words in front of it. Three fixes are applied automatically because every reference project turns them on, and each is reported. A full name sitting in the first name cell with the last name blank is split. A name in ALL CAPS or all lower case in a list that is otherwise title case is title cased, with Mc, Mac, O' and particles handled. Honorifics and suffixes (Dr, Prof, Sir, Jr, MBE and so on) are removed. Swapped first and last names are reported but not changed.

Sorting. Case insensitive, accent folded (Müller sorts with Muller), natural numeric (Table 2 before Table 10). Blanks sort last, numbers before text. Surname particles sort UK style under the particle by default (van der Berg under V) or Dutch style under the main name. A particle is only recognised when it is lower case or when the whole surname is one case, so De Niro stays one name. Mc and Mac are sorted as spelled. Source order is the final tie-break. Numeric text in the table column is read as a number so CSV and Excel sort the same way.

Chunking. Every slide is full except the last. With break on table change, the rows are sorted by table first and a new slide starts whenever the table changes. The first slide of each table uses the LIST BREAK layout with the table in the header. A table that runs past one slide continues on LIST slides. A blank table sorts last, gets its own slide with an empty header and a warning naming the people.

Range. The first and last surname on the slide, Adams to Fisher. A person with no surname shows their name instead.

Rows with no name at all are dropped and the slide that would have held them carries a note, because dropping a row shifts everyone after it.

## Layout

Columns are derived from names per slide, up to 15 names a column and up to four columns. 30 names gives two columns of 15. The row height is the space available divided by the rows per column and the font size follows it, 14 pt at 30 a slide. The table box and the extra field box are sized from the longest value in the data, and the first name box when first and last are separate. The table number is right aligned at the edge of its column. Row boxes do not wrap, a name too long for its box runs over the edge where it can be seen rather than folding onto a second line. Title top left, header below it, range top right, page bottom right.

All of this sits on the layout. Nothing is positioned or formatted on the slides.

## Checking in PowerPoint

Open the deck, View, Slide Master, click the LIST layout, drag a name box, change its font or colour, come out of master view. Every slide should follow. Do the same on LIST BREAK for the header. If a box on a slide does not follow, that slide has a shape with its own position, which this writer never produces, so it would mean PowerPoint rewrote it on open.

The Selection Pane on a slide shows the shape names, row.1.name and so on.

## Differences from the Python reference

- The header for a non-numeric table reads as the value itself (TBC, Top table). The reference writes Table TBC.
- A range end for someone with no surname shows their first name. The reference leaves it blank.
- Numeric text in the table column is read as a number for both CSV and Excel. The reference keeps CSV values as text, so 12A would sort between 12 and 13 from a CSV and after every number from Excel.
- No speaker notes are written. The reference writes the rows, break value, range and page into each slide's notes.
- No build report file. The warnings are on the page.

Everything else, the sort order, the chunk boundaries, the layouts chosen and the warnings, matches the reference decks table_alpha.pptx and table_by_table.pptx slide for slide.

## Tests

`node --test tests/list-builder.test.js` runs the sort, particle, title case, header detection, CSV and chunking cases copied from the Python tests, plus a build that checks the pptx structure. `tests/list-builder.html` runs the same cases in a browser.

## Files

- `list-builder.html`, the page
- `list-builder.js`, the engine and the pptx writer, loads in the browser and in Node
- `tests/list-builder.test.js` and `tests/list-builder.html`

Libraries from CDN, SheetJS 0.18.5 for .xlsx and JSZip 3.10.1 for the package.
