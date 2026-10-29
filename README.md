# JarvisDoc

A desktop document editor for real **Word (`.docx`)** and **Markdown (`.md`)** files, for macOS, Windows, and Linux.

Word files stay Office Open XML. Opening a `.docx` keeps its page size, margins, headers, footers, and embedded images in the package. Fonts, colors, tables, lists, and pictures in the body round-trip back into the same document. Headers and footers are shown on the page and left untouched so page numbers and letterhead survive a save.

Saving an unedited Word file writes that same `.docx` back out, so the original formatting stays intact. **Save as PDF** builds the PDF from that Word file. Markdown keeps headings, lists, and links. Fonts, colors, tables, and images are stored when you save as Word.

## Markdown, learned by using it

Markdown syntax works as you type, in Word documents too. Start a line with `# `, `## `, `- `, `1. ` or `> ` and the line becomes a heading, a list or a quote. Type `**bold**`, `*italic*`, `~~struck~~`, `` `code` `` or `[text](url)` and it formats the moment you close the marker. `---` on its own line draws a divider and ` ``` ` starts a code block. None of this changes the file format — a `.docx` stays a `.docx`.

The status bar says whether the open document would survive as Markdown. **View → Markdown View** (⌥⌘M) shows the live Markdown source, an itemised list of anything that would not come across, and a cheat sheet. When a document is Markdown-ready and you are still in Word, JarvisDoc offers to make the switch; when it is not, saving as Markdown lists exactly what would be dropped first and lets you keep editing the Word file instead.

A small paper character sits in the bottom-right corner. Clicking it opens a short, numbered case for Markdown — one argument at a time, covering permanence, plain text, real diffs, honest headings, searchability, and the fact that nobody owns the format. Where it can, it argues from the document in front of you: on a Word file it will tell you how much of the file on disk is your writing and how much is scaffolding.

It is built to be the opposite of the assistants it resembles. It never speaks while you are typing, waits a minute and a half after launch and half a minute after a document opens, says at most two things in a sitting, never repeats a lesson, and never appears alongside another hint. Close three lessons without reading them and it stops volunteering for good — the character stays, so the whole argument is still one click away. **Help → Why Markdown?** opens it on demand and **View → Markdown Tips** removes it entirely.

Markdown files keep their images: pictures are written to a `<name>.assets` folder beside the `.md` and read back in when you open it. Tables, headings, lists, links and emphasis all round-trip. Occasional hints point out the Markdown equivalent of a toolbar button; each appears at most three times and **View → Markdown Tips** turns them off for good.

## Editing

Right-click anywhere in the page for spelling suggestions, cut/copy/paste, formatting, and — inside a table — row and column commands. Tables are also editable from the **Insert** menu; Tab moves between cells and adds a row at the end of the table, indents in a list, and inserts a real Word tab stop in ordinary text.

**Find and Replace** is ⌥⌘F, plain **Find** is ⌘F. **File → Open Recent** lists recent documents and **Revert to Saved** discards unsaved changes after snapshotting them. Dropping a `.docx`, `.md`, or image file onto the window opens it or inserts it. The window remembers its size and position.

## Version history

JarvisDoc keeps a copy of the document every time you save, plus a draft every five minutes while there are unsaved changes. Open **View → Version History** (⌥⌘V) or click the version count in the status bar to browse them, grouped by day.

Each version can be viewed read-only in place, restored, exported to a separate file, or deleted. **File → Save a Version…** (⌥⌘S) stores a named checkpoint that is never pruned automatically. Restoring loads the old content into the editor and leaves the document unsaved, and any unsaved work is snapshotted first, so a restore can never lose text.

Versions are stored outside the document in the app's support folder and are keyed to the file path, so a copy saved under a new name starts its own history. Identical saves are not duplicated; routine versions are pruned to the most recent 40.

## Printing

**File → Print…** (⌘P / Ctrl+P) sends the page to the system print dialog at the document's own paper size, so a Letter document prints as Letter and an A4 one as A4. Backgrounds and table rules are printed; the toolbar, status bar, find bar, hint bar, Markdown view and version history are not — only the page itself reaches the paper.

**File → Save as PDF** is a separate path that aims for Word fidelity rather than a screenshot of the editor. It converts the saved `.docx` with Pages on macOS, Word on Windows, or LibreOffice on any platform, whichever is installed first. If none is available it falls back to rendering the page itself, which still produces a correct PDF at the right paper size.

## Platform support

| | Minimum version | Notes |
| --- | --- | --- |
| macOS | 11 Big Sur | Universal binary — Intel and Apple Silicon |
| Windows | 10, 11, Server 2016+ | x64 and arm64 |
| Linux | glibc 2.25+ | Ubuntu 18.04+, Debian 10+, RHEL/Fedora 8+ |

The floor is set by the bundled Electron runtime, not by JarvisDoc itself. Electron 37 is the newest release that still supports macOS 11; moving further forward would require macOS 12, and macOS 13 from Electron 45 on.

Window chrome follows the platform: macOS gets the inset traffic lights, Windows a custom title bar with a native overlay for the caption buttons, and Linux keeps its own window frame so the desktop's decorations apply.

Fonts fall back by metric compatibility rather than by name, so a document set in Calibri or Cambria lays out at the same width on a machine that only has Carlito and Caladea. On Debian and Ubuntu those come from `fonts-crosextra-carlito` and `fonts-crosextra-caladea`, which the `.deb` recommends.

## Building

```sh
npm install
npm start                 # run from source

npm run dist:mac          # universal .dmg and .zip
npm run dist:win          # NSIS installer and portable .exe, x64 + arm64
npm run dist:linux        # AppImage, .deb and .rpm, x64 + arm64
```

Each target has to be built on, or cross-built for, its own platform: the Windows installer needs Windows or Wine, and the `.deb` and `.rpm` need `dpkg` and `rpmbuild`. The unpacked trees (`electron-builder --win dir`, `--linux dir`) do build on macOS and are enough to check packaging.
