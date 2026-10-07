# PowerPoint to PDF

Converts presentations to PDF in batches, using the PowerPoint installed on the user's own computer. Every deck is checked before conversion and every PDF is checked afterwards. Whatever could not be checked is reported as not checked. No presentation, font, preview, PDF or report leaves the computer.

Lives at `/ppt-pdf.html`. The homepage links to it under Other Tools and passes on the pixel size from the calculator, which fills in the custom page size.

## Why it needs a helper

A web page cannot start or control PowerPoint. Nothing in a browser allows it. So conversion is done by a small helper script that the user starts, and the page talks to that.

Ways of connecting the two that were looked at

| Route | Verdict |
| --- | --- |
| Convert in the browser with a library | Rejected. It would not be PowerPoint drawing the slides, and fidelity is the point |
| A paid conversion service | Rejected. Files would leave the computer |
| An Office add-in | Cannot open other files or export them in a batch |
| A browser extension with native messaging | Needs two installs, per browser |
| A custom URL handler | Needs an installer to register it |
| A helper with a local web server on 127.0.0.1 | Works, but since Chrome 142 a public site reaching the local machine raises a permission prompt, the open port has to be defended against other sites, and stock macOS has nothing to serve HTTP with |
| A helper that shares a folder with the page | Chosen |

The page already has to write PDFs and reports into an output folder. The helper lives in a `_slidesize` folder inside it and the two exchange small JSON files there. No port is opened, no extra permission is asked for, the same protocol serves Windows and macOS, and because all state is on disk an interrupted batch can be picked up later. The cost is that the page needs the File System Access API, so it runs in Chrome and Edge and not in Safari or Firefox. The page says so when it finds itself somewhere else.

Nothing is installed. The page writes the helper into the folder and shows how to start it.

| System | Helper | Started by |
| --- | --- | --- |
| Windows | `Start-SlideSize-Helper.cmd` and `slidesize-helper.ps1`. Windows PowerShell 5.1, which every Windows 10 and 11 PC has, driving PowerPoint over COM | Double-click the `.cmd` |
| macOS | `Start-SlideSize-Helper.command`. A shell script driving PowerPoint with AppleScript | Terminal. Type `sh` and a space, drag the file in, press Return |

macOS will not run a script that a browser wrote when it is double-clicked, which is why the Mac helper is started from Terminal. If a browser or a security tool refuses to let the page write a script, the Download the helper button gives the same files as a zip.

## How a batch runs

1. Select files. Drop files or folders, or use the buttons.
2. Check presentations. Each file is read in the browser.
3. Choose settings and an output folder. Start the helper.
4. Convert.
5. Review results and warnings.
6. Save or open PDFs and reports.

One presentation is in hand at a time.

- The page copies the deck into `_slidesize/in` as a stream and writes a ticket into `_slidesize/queue`.
- The helper opens the copy in PowerPoint, read only and without a window on Windows, exports the PDF and, if the fidelity check is on, a picture of each slide. It closes the copy without saving and deletes it.
- The helper writes a result into `_slidesize/done`.
- The page opens the PDF, checks it, makes any edits, compares it with PowerPoint's pictures, writes the batch record and the three report files, and removes the pictures it no longer needs.

At most two decks are copied or queued ahead of PowerPoint. In a 120 deck test the page's memory grew by about 8 KB a file.

The original files are never opened by PowerPoint and never written to. Because the copy has a name of its own, a presentation the user already has open is never the one the helper closes.

The helper is two processes. A watchdog hands out tickets and keeps time. A worker talks to PowerPoint. If a file runs past its time limit the watchdog stops the worker, records the failure and starts a new worker for the next file. It will only close PowerPoint itself when the helper started it and it has no windows of the user's open.

### Interrupting and resuming

The record of the batch is `_slidesize/batch.json`, rewritten after every file. Closing the tab, the browser or the computer loses nothing that had finished. On return the page offers to resume. Finished files stay finished and files that were in hand are converted again.

Browsers do not let a page keep reading files after it has been closed unless the user allowed it on every visit. Where access has lapsed the page asks for the files or their folder to be selected again and matches them by name, size and date. Adding a folder in the first place means one confirmation covers every file.

## What is checked before conversion

A `.pptx` is a zip. Only its directory and XML parts are read, by slicing the file, so a 2 GB deck costs a few megabytes and its media is never loaded.

- Damaged, password protected, empty and non-PowerPoint files
- Slide count, slide size, hidden slides
- Fonts named on slides, layouts, masters and the theme, and fonts embedded in the deck
- Video and audio, embedded or linked, with the position and poster frame of each video. A poster frame that is one flat colour is found by looking at the picture
- Pictures and media linked to a file path, the internet or a network share
- Animations, with exit effects and overlapping animated shapes picked out, because a PDF shows everything at once
- Transitions including Morph, ActiveX controls, 3D models, embedded objects
- Speaker notes, slide number fields, a password to modify, macros

A legacy `.ppt` gives up only its slide size, slide count and whether it is encrypted. The rest is reported by PowerPoint after it has opened the file.

A deck that links to the internet or a network share is held back until Allow linked files from the network is turned on, because PowerPoint fetches linked files as soon as it opens a deck.

## What is checked after conversion

- The PDF opens
- Page count against the slides, hidden slides and range
- Page size against the chosen size
- Fonts read back from the PDF. A missing font, a font whose licence forbids embedding and a font that is named but not embedded are reported as three different things
- PDF/A, see below
- Each PDF page against PowerPoint's own picture of the slide

### The picture comparison

Each compared page is drawn in the browser with pdf.js and set against the picture PowerPoint saved of the same slide, both at 384 pixels wide and cut into 8 pixel blocks. A block is flagged when its average brightness differs or when its pixels differ a lot on average. Resized pages are lined up first, so a letterboxed or cropped page is compared with the right part of the slide.

It finds content that is missing, moved or reflowed. It cannot prove a page is right. The two programs smooth edges, gradients and shadows differently, so small differences are ignored on purpose. A flagged page is marked suspected. A clean comparison is evidence and is never shown as a guarantee. It is not run on notes pages.

### PDF/A

PowerPoint has one PDF/A setting, ISO 19005-1. The part and level it actually wrote are read back from each file.

| Shown | When |
| --- | --- |
| Passed | veraPDF is installed on the computer and has checked the finished file |
| Failed | veraPDF says so, or one of the structural rules checked here is plainly broken, such as a font that is not embedded or no PDF/A declaration at all |
| Not verified | Everything else, including a file that declares PDF/A and breaks none of the rules checked here |

Choosing PDF/A is never treated as proof. veraPDF is free, from verapdf.org. If it is on the path or in its default folder the Windows helper finds and runs it. When the page has edited the file, by resizing for instance, the validator is asked about the edited file.

## Resizing

Pages are resized after PowerPoint has made the PDF, by putting one scale and offset in front of each page's content and changing the page box around it. The scale is the same on both axes, so nothing can be stretched. Text stays text, vectors stay vectors, and tags and links are kept, with link areas moved to match. The presentation is not touched.

Fit adds margins, which can be left unpainted or painted white or black. Fill crops, and the cropped strips are flagged for review with how much is lost. Removing links and stripping metadata take the data out of the file, not just out of sight.

## Results

| State | Meaning |
| --- | --- |
| Completed | A PDF, and nothing to report beyond notes |
| Completed with warnings | A PDF with warnings, or with review items that have been marked as reviewed |
| Needs review | A PDF with a confirmed error or something that needs a person to look, not yet marked as reviewed |
| Failed | No usable PDF |

A PDF that exists can always be opened and saved whatever its state. Marking a slide as reviewed changes the state and removes no warning.

Every issue records the file, the slide number, the PDF page where there is one, a severity, a description, the likely impact and a suggested fix, and one of three levels of certainty.

- Confirmed. Read from the file or reported by PowerPoint
- Suspected. Inferred, likely but not proven
- Not verified. A check that could not be run

The reports are `SlideSize report <date>.html`, `.csv` and `.json` in the output folder, rewritten after every file. They record settings, engine and version, processing time, excluded slides, font substitutions, rasterised text, media handling, the fidelity check and PDF/A.

## Platform differences

| Setting | Windows | macOS |
| --- | --- | --- |
| Standard PDF, original size | Yes | Yes |
| Page size presets, fit or fill | Yes | Yes |
| Hidden slides in or out | Yes | Yes, see below |
| Slide range | Yes | Yes, see below |
| Fidelity check | Yes | Yes |
| Video placeholders | Yes | Yes |
| Remove links, strip metadata | Yes | Yes |
| PDF/A | Yes | No |
| Notes pages | Yes | No |
| Image compression choice | Yes | No |
| Bitmap text when fonts cannot be embedded | Yes | No |
| Accessibility tags, document properties, comments and ink | Yes | No |

PowerPoint for Mac has no export options a script can reach, so those settings are switched off on a Mac with the reason beside each. Once the helper is running it reports what the installed PowerPoint can do and that replaces the page's guess.

To leave slides out on macOS the helper first hides them in its temporary copy, which changes nothing else. If the PDF then has too many pages, because that version of PowerPoint exports hidden slides anyway, it removes the unwanted slides from the copy and exports again. Slide numbers printed on later slides then count from the slides that remain, and the report says so.

## Limits worth knowing

- Chrome or Edge on a desktop computer.
- A PC managed to block PowerShell scripts cannot run the Windows helper.
- On macOS PowerPoint shows each presentation briefly while it converts it.
- While the user has presentations of their own open on Windows, a file that stalls PowerPoint cannot be cleared automatically, because the helper will not close a PowerPoint that is showing their work. The batch waits and says why.
- A PDF above 300 MB is not opened for checking or resizing. It is kept as PowerPoint wrote it and reported as not verified.
- Decks above 2 GB are outside what has been tested.
- Image compression is PowerPoint's own two choices. There is no finer control because PowerPoint offers none.
- The tab must stay open. It can be in the background.

## What has been tested, and what has not

Tested here, on Linux, with no PowerPoint

- The page, end to end in Chromium, with LibreOffice standing in for PowerPoint and a private browser folder standing in for a real one. Checking, settings, naming, converting, resizing, the comparison, review, reports, failures, retry, skip, stop, resume, PDF/A states, both platforms' capability sets, and a 120 deck batch for memory and disk.
- The Windows helper under PowerShell 7 and the macOS helper under `sh`, each with PowerPoint replaced by a stand-in. The queue, time limits, cancelling, recovery, resuming, one helper per folder, refusal of tickets that point outside the folder, and both ways of leaving slides out on macOS.
- The deck reader, the PDF reader and editor, geometry, names, issues, states and reports, as unit tests.

Not tested here

- The Windows helper driving real PowerPoint. The COM calls are written from Microsoft's documentation and have not been run.
- The Windows helper under Windows PowerShell 5.1. It is written to 5.1's syntax and checked for anything newer, but was run under PowerShell 7.
- How faithful PowerPoint's PDFs are. That is PowerPoint's doing and can only be judged on real decks.
- Whether Chrome lets the page write the helper scripts into a folder on every machine. The zip download is there for when it does not.

See the end of this file for the macOS result once it has been run on a Mac with PowerPoint.

## Files

| File | What it is |
| --- | --- |
| `ppt-pdf.html` | The page |
| `pptpdf-app.js` | Page logic and interface |
| `pptpdf-core.js` | Settings, platform capability, geometry, names, tickets, issues, states, reports, the picture comparison. No DOM, runs in Node |
| `pptpdf-inspect.js` | Reads `.pptx` and `.ppt` files. No DOM, runs in Node |
| `pptpdf-pdf.js` | Reads back and edits the finished PDF. No DOM, runs in Node |
| `pdf-lib.min.js` | pdf-lib 1.17.1 by Andrew Dillon, MIT licence, vendored |
| `helper/` | The helper scripts the page writes into the output folder |
| `tests/` | Not needed on the site |

pdf.js is loaded from cdnjs when a page first has to be drawn, as PDF to Slides does.

## Tests

From the repository root

```
node --test
```

The unit tests have no dependencies. The helper tests need a POSIX `sh` for the macOS helper and PowerShell for the Windows one. Set `PWSH` to its path if it is not called `pwsh` or `powershell`. Without it those are skipped.

`tests/pptpdf/browser/` holds the checks that need a real browser. See the note at the top of `run.js`.

`tests/pptpdf/make_decks.py` rebuilds the test presentations and PDFs.
