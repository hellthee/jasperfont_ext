# jasperfont_ext

JasperReports font extension for Simplified Chinese. Renders CJK text
correctly in PDF export and keeps line heights consistent between PDF and
`.docx` export.

Built and tested against **JasperReports 6.6.0** (iText 2.1.7.js6).

## Install

Put `dist/jasperreports-ext-fonts.jar` on the classpath:

| Target | Location |
|---|---|
| Standalone app | `lib/` |
| Web app | `WEB-INF/lib/` |
| JasperReports Server | `<server>/WEB-INF/lib/` then restart |
| iReport / NetBeans designer | Tools → Options → iReport → Classpath → Add JAR, then restart |

JasperReports discovers the extension automatically via
`jasperreports_extension.properties`. No code changes needed to register it.

The designer classpath and the runtime classpath are separate. Adding the
JAR to iReport only affects the designer — the application server needs it
too.

## Use

```xml
<font fontName="Chinese" size="12" pdfEncoding="Identity-H" isPdfEmbedded="true"/>
```

Two families are registered, both backed by the same files:

- `Chinese` — short alias
- `Noto Sans SC Jasper` — full name

`pdfEncoding="Identity-H"` is mandatory. A single-byte encoding cannot
address CJK code points; characters outside it render blank with no error.
`fonts.xml` sets it at the family level, but setting it per-element too is
harmless and makes the requirement explicit in the report.

Report-wide default:

```xml
<property name="net.sf.jasperreports.default.font.name" value="Chinese"/>
```

See `examples/ChineseReportDemo.jrxml` and
`java/ChinesePdfExportExample.java`.

## Making .docx match PDF

Font metrics alone do not make Word agree with the PDF. Two separate
defects in the JasperReports `.docx` exporter need patching after export.

### Line height

iText derives line height from the font's `hhea` table: for this font
(ascent 1160, descent -288, UPM 1000) that is 1.448 × font size — 17.376pt
at 12pt. The exporter writes `<w:spacing w:lineRule="auto" w:line="240"/>`
on every paragraph, which pins Word to exactly 12pt regardless of font
size and overrides the font's own metrics.

`java/DocxLineSpacingFixer.java` rewrites each paragraph's spacing to
`lineRule="exact"` with the iText-equivalent value, derived from that
paragraph's font size. Image paragraphs are left alone.

### Page breaks

The exporter emits each report page as a separate `<w:tbl>` and puts
`<w:pageBreakBefore/>` in the first cell of every table after the first.
Word then starts a new page even when the content fits on the current one,
producing trailing blank pages.

`java/DocxPageBreakRemover.java` strips those elements.

### Order

```java
JasperExportManager.exportReportToDocxFile(jasperPrint, "/tmp/report.docx");
DocxPageBreakRemover.fixInPlace("/tmp/report.docx");   // first
DocxLineSpacingFixer.fixInPlace("/tmp/report.docx");   // second
```

Page breaks first — no point computing spacing for paragraphs on a page
that is about to be removed. Both classes are dependency-free (Java 8+,
`java.util.zip` and `java.util.regex` only) and can be run from the
command line.

## Building

```bash
pip install fonttools
python3 build.py                         # downloads the source TTCs
python3 build.py --ttc-dir /path/to/ttc  # or use local ones
python3 build.py --family "Your Name"    # override the family name
```

`build.py` checks common system font directories first and only downloads
if the TTCs are not found. It verifies its own output and fails loudly
rather than shipping a silently broken font.

### What the build does, and why

**Extract** face index 2 (`Noto Sans CJK SC`) from `NotoSansCJK-*.ttc`.

**Convert CFF outlines to TrueType.** This is the step that makes the font
work at all. noto-cjk ships CFF-flavoured OpenType (sfnt magic `OTTO`).
iText 2.1.7 cannot embed CFF OpenType with Identity-H — it raises no
error and the glyphs simply do not appear in the PDF. Conversion is via
cu2qu at 1.0 em-unit tolerance (visually lossless), producing `glyf` +
`loca` and sfnt magic `0x00010000`.

Because the font has 65535 glyphs, `post` is written as format 3.0 —
format 2.0's glyph-name index is a uint16 and overflows above 32767
glyphs. iText and JasperReports address glyphs by cmap, not by name.

**Align OS/2 metrics.** Upstream `sTypoAscender`/`sTypoDescender` are
880/−120, which gives Word 12.000pt at 12pt against iText's 17.376pt.
The build sets them to `hhea`'s 1160/−288, mirrors them into
`usWinAscent`/`usWinDescent`, and sets `USE_TYPO_METRICS` — which requires
upgrading OS/2 to version 4.

**Rename the family** — see below.

## License

Extension code, `build.py`, the Java utilities and the XML: MIT.

The bundled font derives from Noto Sans CJK SC, licensed under the
[SIL Open Font License 1.1](OFL.txt). The build modifies it (outline
conversion and metric changes), and OFL sections 3 and 5 prohibit
distributing a modified font under its original name — so the family is
renamed to `Noto Sans SC Jasper` in both the `name` table and `fonts.xml`.
If you fork this and change the font further, rename it again; do not
reintroduce a Noto reserved font name.
