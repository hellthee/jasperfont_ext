package com.example.jasper;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * OoxmlFontSubstituter
 *
 * Replaces the extension's Chinese font in a JasperReports .docx or .pptx
 * export with system fonts: one for Latin text, one for East Asian text.
 * Intended for recipients who do not have Noto Sans CJK SC installed.
 *
 * WHY A POST-PROCESSOR:
 *   fonts.xml exportFonts holds one font name per export format, for the
 *   whole JAR. JasperReports 6.6.0 writes that single name into every
 *   script slot, e.g. in docx:
 *     <w:rFonts w:ascii="X" w:hAnsi="X" w:eastAsia="X" w:cs="X"/>
 *   Both formats support a separate font per script; JasperReports just
 *   does not use it. This class fills the slots separately.
 *
 * WHAT IS REWRITTEN:
 *   Only fonts named "Noto Sans CJK SC", "Noto Sans SC" or "Chinese" (the
 *   names this extension has written, across versions). Other fonts are
 *   left alone.
 *     docx  every word/*.xml part: w:rFonts
 *           ascii, hAnsi, cs -> Latin font; eastAsia -> East Asian font.
 *           Other attributes (e.g. w:hint) are kept.
 *     pptx  every ppt/*.xml part: a:latin, a:ea, a:cs in run properties.
 *           A missing a:ea or a:cs after a:latin is added, so East Asian
 *           text does not fall through to the theme font.
 *
 * DO NOT combine with DocxLineSpacingFixer. That class sets exact line
 * heights computed from Noto Sans CJK SC metrics, which are wrong for
 * Arial and Microsoft YaHei. Leave JasperReports' single line spacing in
 * place so Word uses the substituted fonts' own metrics.
 *
 * USAGE:
 *   DocxPageBreakRemover.fixInPlace(path);      // docx only
 *   OoxmlFontSubstituter.fixInPlace(path);      // Arial + Microsoft YaHei
 *   OoxmlFontSubstituter.fixInPlace(path, "Arial", "SimSun");
 *
 * REQUIREMENTS:
 *   Java 8+, no additional dependencies.
 */
public class OoxmlFontSubstituter {

    public static final String DEFAULT_LATIN_FONT      = "Arial";
    public static final String DEFAULT_EAST_ASIAN_FONT = "Microsoft YaHei";

    /** Font names this extension has written into exports. */
    private static final Pattern SOURCE_FONT = Pattern.compile(
            "Noto Sans CJK SC|Noto Sans SC|Chinese");

    private static final Pattern DOCX_RFONTS = Pattern.compile(
            "<w:rFonts\\b([^>]*?)/>");
    private static final Pattern DOCX_SLOT = Pattern.compile(
            "\\s+w:(ascii|hAnsi|eastAsia|cs)=\"([^\"]*)\"");

    /** a:latin, optionally followed by a:ea and a:cs (schema order). */
    private static final Pattern PPTX_LATIN_GROUP = Pattern.compile(
            "<a:latin\\b[^>]*?typeface=\"([^\"]*)\"[^>]*/>"
            + "(\\s*<a:ea\\b[^>]*/>)?(\\s*<a:cs\\b[^>]*/>)?");
    /** a:ea or a:cs not preceded by a substituted a:latin. */
    private static final Pattern PPTX_EA_CS = Pattern.compile(
            "<a:(ea|cs)\\b[^>]*?typeface=\"([^\"]*)\"[^>]*/>");

    /** Substitute with Arial + Microsoft YaHei. */
    public static void fix(String inputPath, String outputPath) throws IOException {
        fix(inputPath, outputPath, DEFAULT_LATIN_FONT, DEFAULT_EAST_ASIAN_FONT);
    }

    /**
     * Substitute with the given fonts.
     *
     * @param inputPath     JasperReports .docx or .pptx export
     * @param outputPath    where to write the result
     * @param latinFont     font for Latin (and complex-script) text
     * @param eastAsianFont font for East Asian text
     */
    public static void fix(String inputPath, String outputPath,
                           String latinFont, String eastAsianFont) throws IOException {
        File outputFile = new File(outputPath).getAbsoluteFile();
        File tempFile   = File.createTempFile("ooxml_fs_", ".tmp", outputFile.getParentFile());
        int substituted = 0;

        String docxFonts = "<w:rFonts w:ascii=\"" + latinFont + "\" w:hAnsi=\"" + latinFont
                + "\" w:eastAsia=\"" + eastAsianFont + "\" w:cs=\"" + latinFont + "\"";
        String pptxFonts = "<a:latin typeface=\"" + latinFont + "\"/>"
                + "<a:ea typeface=\"" + eastAsianFont + "\"/>"
                + "<a:cs typeface=\"" + latinFont + "\"/>";

        try (ZipInputStream  zin  = new ZipInputStream(new FileInputStream(inputPath));
             ZipOutputStream zout = new ZipOutputStream(new FileOutputStream(tempFile))) {

            byte[] chunk = new byte[8192];
            ZipEntry entry;
            while ((entry = zin.getNextEntry()) != null) {
                String name = entry.getName();

                ByteArrayOutputStream buf = new ByteArrayOutputStream();
                int n;
                while ((n = zin.read(chunk)) != -1) {
                    buf.write(chunk, 0, n);
                }
                byte[] data = buf.toByteArray();

                boolean docxPart = name.startsWith("word/") && name.endsWith(".xml");
                boolean pptxPart = name.startsWith("ppt/")  && name.endsWith(".xml");

                if (docxPart) {
                    String xml = new String(data, StandardCharsets.UTF_8);
                    StringBuffer sb = new StringBuffer(xml.length());
                    Matcher m = DOCX_RFONTS.matcher(xml);
                    while (m.find()) {
                        String attrs = m.group(1);
                        boolean hit = false;
                        Matcher slot = DOCX_SLOT.matcher(attrs);
                        while (slot.find()) {
                            if (SOURCE_FONT.matcher(slot.group(2)).matches()) {
                                hit = true;
                                break;
                            }
                        }
                        String replacement = m.group();
                        if (hit) {
                            // keep non-slot attributes such as w:hint
                            String kept = DOCX_SLOT.matcher(attrs).replaceAll("");
                            replacement = docxFonts + kept + "/>";
                            substituted++;
                        }
                        m.appendReplacement(sb, Matcher.quoteReplacement(replacement));
                    }
                    m.appendTail(sb);
                    data = sb.toString().getBytes(StandardCharsets.UTF_8);

                } else if (pptxPart) {
                    String xml = new String(data, StandardCharsets.UTF_8);

                    StringBuffer sb = new StringBuffer(xml.length());
                    Matcher m = PPTX_LATIN_GROUP.matcher(xml);
                    while (m.find()) {
                        String replacement = m.group();
                        if (SOURCE_FONT.matcher(m.group(1)).matches()) {
                            replacement = pptxFonts;
                            substituted++;
                        }
                        m.appendReplacement(sb, Matcher.quoteReplacement(replacement));
                    }
                    m.appendTail(sb);

                    // a:ea / a:cs naming a source font without one in a:latin
                    String pass1 = sb.toString();
                    sb = new StringBuffer(pass1.length());
                    m = PPTX_EA_CS.matcher(pass1);
                    while (m.find()) {
                        String replacement = m.group();
                        if (SOURCE_FONT.matcher(m.group(2)).matches()) {
                            String font = "ea".equals(m.group(1)) ? eastAsianFont : latinFont;
                            replacement = "<a:" + m.group(1) + " typeface=\"" + font + "\"/>";
                            substituted++;
                        }
                        m.appendReplacement(sb, Matcher.quoteReplacement(replacement));
                    }
                    m.appendTail(sb);
                    data = sb.toString().getBytes(StandardCharsets.UTF_8);
                }

                zout.putNextEntry(new ZipEntry(name));
                zout.write(data);
                zout.closeEntry();
                zin.closeEntry();
            }
        } catch (IOException e) {
            tempFile.delete();
            throw e;
        }

        Files.move(tempFile.toPath(), outputFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
        System.out.printf("[OoxmlFontSubstituter] %d font reference(s) -> %s / %s%n",
                substituted, latinFont, eastAsianFont);
    }

    /** Substitute in place with Arial + Microsoft YaHei. */
    public static void fixInPlace(String filePath) throws IOException {
        fix(filePath, filePath, DEFAULT_LATIN_FONT, DEFAULT_EAST_ASIAN_FONT);
    }

    /** Substitute in place with the given fonts. */
    public static void fixInPlace(String filePath, String latinFont, String eastAsianFont)
            throws IOException {
        fix(filePath, filePath, latinFont, eastAsianFont);
    }

    public static void main(String[] args) throws IOException {
        if (args.length < 1) {
            System.err.println("Usage: OoxmlFontSubstituter <input.docx|.pptx> [output]"
                    + " [latinFont] [eastAsianFont]");
            System.err.println("       Output omitted: fixes in place."
                    + " Fonts default to Arial / Microsoft YaHei.");
            System.exit(1);
        }
        String output  = args.length > 1 ? args[1] : args[0];
        String latin   = args.length > 2 ? args[2] : DEFAULT_LATIN_FONT;
        String eastAsn = args.length > 3 ? args[3] : DEFAULT_EAST_ASIAN_FONT;
        fix(args[0], output, latin, eastAsn);
    }
}
