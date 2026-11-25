# jasperfont_ext

JasperReports font extension for Simplified Chinese. Renders CJK text
correctly in PDF export and keeps line heights consistent between PDF and
`.docx` export.

Supports **JasperReports 6.6.0** (PDF via iText 2.1.7.js6) and
**6.21.5** (PDF via OpenPDF 1.3.32). See [Upgrading to 6.21.5](#upgrading-to-6215).

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
- `Noto Sans SC` — full name

`pdfEncoding="Identity-H"` is mandatory. A single-byte encoding cannot
address CJK code points; characters outside it render blank with no error.
`fonts.xml` sets it at the family level, but setting it per-element too is
harmless and makes the requirement explicit in the report.

### Set the default font application-wide

Text elements without an explicit `fontName` use JasperReports' default,
`SansSerif`, which is not in this extension. In PDF, Chinese text in such
an element renders blank; in Office exports the document names a font no
machine has. Set the default once in your application's
`jasperreports.properties`:

```properties
net.sf.jasperreports.default.font.name=Chinese
```

A per-report `<property name="net.sf.jasperreports.default.font.name"
value="Chinese"/>` also works, but the properties file covers every report
and every export format.

See `examples/ChineseReportDemo.jrxml` and
`java/ChinesePdfExportExample.java`.

## Office exports

Office formats reference fonts by name rather than embedding them, so what
the recipient sees depends on which fonts their machine has. `fonts.xml`
maps the export names:

| Format | Font name written | Why |
|---|---|---|
| docx, pptx, rtf, odt, ods | `Noto Sans CJK SC` | Internal family name of the bundled TTF — the name Office finds once that TTF is installed. `Noto Sans SC` would not match it. |
| xlsx, xls, metadata xlsx | `Arial` | A cell font has no per-script slot; Excel's own fallback renders the Chinese. `JRXlsxMetadataExporter` reads its own key, `net.sf.jasperreports.metadata.xlsx`. |

### Two tiers of recipients

**Tier 1 — Noto Sans CJK SC installed.** Install both TTFs on the
recipient's machine: `python3 build.py --keep-ttf` leaves them in `dist/`.
Install those, not Google Fonts' "Noto Sans SC": that is a different
build with a different name. Word then renders with the same font as the
PDF and the same metrics: it reads the win metrics (1160 + 288), giving
1.448 × font size, which is what iText uses.

**Tier 2 — no Noto.** `java/OoxmlFontSubstituter.java` rewrites docx and
pptx so Latin text uses Arial and East Asian text uses Microsoft YaHei.
fonts.xml can't do this: `exportFonts` holds one name per format, and
JasperReports writes that one name into every script slot. Both formats
support a font per script. The substituter fills them separately and
leaves fonts it didn't write alone.

```java
JasperExportManager.exportReportToDocxFile(jasperPrint, path);
DocxPageBreakRemover.fixInPlace(path);              // every docx
if (!recipientHasNoto) {
    OoxmlFontSubstituter.fixInPlace(path);          // Arial + Microsoft YaHei
}
```

pptx: call `OoxmlFontSubstituter.fixInPlace(path)` for tier 2. It needs
no page-break step.

In tier 2 the output does not match the PDF exactly, since it uses
different fonts. That is the tradeoff.

### Page breaks

The docx exporter writes each report page as a separate `<w:tbl>` and
puts `<w:pageBreakBefore/>` in the first cell of every page after the
first. These are real page boundaries — the PDF has the same pages — and
must stay. Removing them runs separate report pages together in Word.

`java/DocxPageBreakRemover.java` removes a break only when the page it
starts is blank: no visible text and no image. That removes the blank pages
Word can show, and keeps every page that has content. Run it for both
tiers.

### Line height

The exporter writes `<w:spacing w:lineRule="auto" w:line="240"/>`: single
spacing, i.e. 1 × the font's own natural line height. With Noto Sans CJK
SC installed that already equals iText's 1.448 ×, so Word matches the PDF
without help.

`java/DocxLineSpacingFixer.java` converts this to `lineRule="exact"` at
1.448 × font size. That's redundant in tier 1 and **wrong in tier 2**: it
would force Noto's line height onto Arial/YaHei text and can clip glyphs.
Never run it after `OoxmlFontSubstituter`.

All three classes are dependency-free (Java 8+, `java.util.zip` and
`java.util.regex` only) and can be run from the command line.

## Upgrading to 6.21.5

Checked against the 6.21.5 sources. The docx/pptx/xls/xlsx exporters
behave as in 6.6.0 where this project depends on them. They write one
font name into every script slot, write the same page breaks and line
spacing, and use the same exporter keys. The font extension format is
unchanged. No change to the JAR or the post-processors is needed.

What changes:

- **PDF engine is OpenPDF 1.3.32, not iText 2.1.7.** The bundled
  TrueType font works with both. Remove `itext-2.1.7.js*.jar` from the
  classpath: both libraries use the `com.lowagie` packages, so having both
  gives duplicate-class conflicts.
- **Recompile `.jasper` files with 6.21.5.** iReport's last release
  (5.6.0) bundles JasperReports 5.6. It can still edit JRXML, but compile
  at runtime or use Jaspersoft Studio.
- **`isStretchWithOverflow` is deprecated.** Use
  `textAdjust="StretchHeight"`. The old attribute still works.
- **Leave `net.sf.jasperreports.export.docx.embed.fonts` (and
  `pptx.embed.fonts`) off.** It copies the whole TTF from `fonts.xml` into
  every file, with no subsetting. With these fonts that adds tens of MB per
  document.

## Building

```bash
pip install fonttools
python3 build.py                         # downloads the source TTCs
python3 build.py --ttc-dir /path/to/ttc  # or use local ones
python3 build.py --family "Your Name"    # rename the font (optional)
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
[SIL Open Font License 1.1](OFL.txt). The build modifies it — outline
conversion and metric changes — and keeps the upstream name.

That is permitted: Noto declares no Reserved Font Name. Its notice reads
`Copyright 2010-2012 Google Corporation`, with no "with Reserved Font
Name" clause, and OFL sections 3 and 5 restrict only Reserved Font Names.

What OFL section 2 does require of every distribution, modified or not,
is that the copyright notice and license text accompany the font. Both
are in `OFL.txt`, which also records what was modified. Keep that file
alongside the JAR if you redistribute it.

`build.py --family NAME` rewrites the font's internal family name. It is
off by default and exists for forks that alter the outlines or metrics
further and would rather not answer to the upstream name.
