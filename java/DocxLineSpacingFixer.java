package com.example.jasper;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.regex.*;
import java.util.zip.*;

/**
 * DocxLineSpacingFixer
 *
 * Post-processes a JasperReports-exported .docx file to fix Chinese (CJK)
 * character height consistency between PDF and Word exports.
 *
 * PROBLEM:
 *   JasperReports 6.6.0 hardcodes <w:spacing w:lineRule="auto" w:line="240"/>
 *   in exported .docx files. This means Word uses 1x the nominal font size
 *   as line height (e.g. 12pt font → 12pt line). But iText (used for PDF)
 *   uses the font's hhea metrics: for Noto Sans SC, that is 1448/1000 × fontSize
 *   = 17.376pt at 12pt — 44.8% taller than Word's default.
 *
 * FIX:
 *   After JasperReports exports the .docx, call DocxLineSpacingFixer.fix()
 *   on the output file. It rewrites every text paragraph's <w:spacing>
 *   to use lineRule="exact" with the correct line height matching iText.
 *
 * FORMULA:
 *   word_line_value = round((HHEA_CELL / UPM) * fontPt * 20)
 *   where HHEA_CELL = 1448 (Noto Sans SC hhea ascent - descent)
 *         UPM       = 1000 (units per em)
 *         fontPt    = font size in points
 *         × 20      = convert points to twentieths-of-a-point (OOXML unit)
 *
 * USAGE:
 *   // After JasperReports exports:
 *   JasperExportManager.exportReportToDocxFile(print, "/tmp/report.docx");
 *   DocxLineSpacingFixer.fix("/tmp/report.docx", "/tmp/report_fixed.docx");
 *
 *   // Or fix in-place:
 *   DocxLineSpacingFixer.fixInPlace("/tmp/report.docx");
 *
 * REQUIREMENTS:
 *   - Java 8+
 *   - No additional dependencies (uses only java.util.zip and java.util.regex)
 *   - jasperreports-ext-fonts.jar with patched Noto Sans SC (USE_TYPO_METRICS=true)
 */
public class DocxLineSpacingFixer {

    // Noto Sans SC hhea metrics (after font patch: ascent=1160, descent=-288)
    private static final double HHEA_CELL     = 1448.0;  // ascent(1160) - descent(-288)
    private static final double UPM           = 1000.0;
    private static final int    DEFAULT_SZ_HP = 24;       // 12pt fallback (half-points)

    // Regex patterns — pre-compiled for performance
    private static final Pattern PARA_PATTERN    = Pattern.compile(
            "<w:p(?:\\s[^>]*)?>.*?</w:p>", Pattern.DOTALL);
    private static final Pattern SZ_PATTERN      = Pattern.compile(
            "<w:sz w:val=\"(\\d+)\"");
    private static final Pattern SPACING_PATTERN = Pattern.compile(
            "<w:spacing([^/]*?)/>");
    private static final Pattern LINE_ATTR       = Pattern.compile(
            "w:line=\"[^\"]*\"");
    private static final Pattern LINE_RULE_ATTR  = Pattern.compile(
            "w:lineRule=\"[^\"]*\"");
    private static final Pattern PPR_CLOSE       = Pattern.compile(
            "</w:pPr>");

    /**
     * Fix line spacing in a JasperReports .docx export.
     *
     * @param inputPath  path to the JR-exported .docx
     * @param outputPath path for the fixed output .docx
     */
    public static void fix(String inputPath, String outputPath) throws IOException {
        File inputFile  = new File(inputPath);
        File outputFile = new File(outputPath);
        File tempFile   = File.createTempFile("docx_fix_", ".docx", outputFile.getParentFile());

        try (ZipInputStream  zin  = new ZipInputStream(new FileInputStream(inputFile));
             ZipOutputStream zout = new ZipOutputStream(new FileOutputStream(tempFile))) {

            ZipEntry entry;
            while ((entry = zin.getNextEntry()) != null) {
                zout.putNextEntry(new ZipEntry(entry.getName()));

                if ("word/document.xml".equals(entry.getName())) {
                    byte[] xmlBytes = readAllBytes(zin);
                    String xml      = new String(xmlBytes, StandardCharsets.UTF_8);
                    String fixed    = processDocumentXml(xml);
                    zout.write(fixed.getBytes(StandardCharsets.UTF_8));
                } else {
                    copyStream(zin, zout);
                }

                zout.closeEntry();
                zin.closeEntry();
            }
        }

        // Atomic replace
        if (outputFile.exists()) outputFile.delete();
        if (!tempFile.renameTo(outputFile)) {
            copyFile(tempFile, outputFile);
            tempFile.delete();
        }
    }

    /**
     * Fix line spacing in-place (overwrites the original file).
     */
    public static void fixInPlace(String filePath) throws IOException {
        String tempPath = filePath + ".tmp";
        fix(filePath, tempPath);
        new File(filePath).delete();
        new File(tempPath).renameTo(new File(filePath));
    }

    // -------------------------------------------------------------------------
    // Core XML processing
    // -------------------------------------------------------------------------

    private static String processDocumentXml(String xml) {
        StringBuilder result = new StringBuilder(xml.length() + 4096);
        Matcher paraMatcher = PARA_PATTERN.matcher(xml);
        int pos = 0;
        int injected = 0, updated = 0, skipped = 0;

        while (paraMatcher.find()) {
            result.append(xml, pos, paraMatcher.start());
            String para = paraMatcher.group();
            int sz = dominantSz(para);

            if (sz < 0) {
                // No text runs — leave unchanged (image paragraphs etc.)
                result.append(para);
                skipped++;
            } else {
                String fixed = fixParagraph(para, sz);
                result.append(fixed);
                if (!fixed.equals(para)) {
                    if (SPACING_PATTERN.matcher(para).find()) updated++;
                    else injected++;
                }
            }
            pos = paraMatcher.end();
        }
        result.append(xml, pos, xml.length());

        System.out.printf("[DocxLineSpacingFixer] injected=%d  updated=%d  skipped=%d%n",
                injected, updated, skipped);
        return result.toString();
    }

    /**
     * Fix or inject w:spacing in a single paragraph.
     */
    private static String fixParagraph(String para, int szHalfPts) {
        int lineVal = lineValForSz(szHalfPts);
        String newSpacing = String.format(
                "<w:spacing w:lineRule=\"exact\" w:line=\"%d\" w:after=\"0\" w:before=\"0\"/>",
                lineVal);

        Matcher spacingMatcher = SPACING_PATTERN.matcher(para);
        if (spacingMatcher.find()) {
            // Has existing w:spacing — update line and lineRule
            String attrs = spacingMatcher.group(1);
            if (attrs.contains("w:lineRule") || attrs.contains("w:line")) {
                String newAttrs = LINE_ATTR.matcher(attrs)
                        .replaceFirst("w:line=\"" + lineVal + "\"");
                newAttrs = LINE_RULE_ATTR.matcher(newAttrs)
                        .replaceFirst("w:lineRule=\"exact\"");
                // If there was no w:line attr, append it
                if (!newAttrs.contains("w:line=")) {
                    newAttrs += " w:line=\"" + lineVal + "\"";
                }
                if (!newAttrs.contains("w:lineRule=")) {
                    newAttrs += " w:lineRule=\"exact\"";
                }
                return para.replace(spacingMatcher.group(0),
                        "<w:spacing" + newAttrs + "/>");
            }
            return para;
        }

        // No w:spacing — inject before </w:pPr>
        Matcher pprClose = PPR_CLOSE.matcher(para);
        if (pprClose.find()) {
            return para.substring(0, pprClose.start())
                    + newSpacing
                    + para.substring(pprClose.start());
        }

        return para; // no w:pPr, leave unchanged
    }

    /**
     * Find the most common w:sz (half-point) value in this paragraph.
     * Returns -1 if no font size found (paragraph has no text runs).
     */
    private static int dominantSz(String para) {
        Matcher m = SZ_PATTERN.matcher(para);
        int bestVal = -1;
        int bestCount = 0;
        java.util.Map<Integer, Integer> counts = new java.util.HashMap<>();
        while (m.find()) {
            int val = Integer.parseInt(m.group(1));
            int count = counts.merge(val, 1, Integer::sum);
            if (count > bestCount) {
                bestCount = count;
                bestVal = val;
            }
        }
        return bestVal; // -1 if no sz found
    }

    /**
     * Compute the Word w:line value (twentieths-of-a-point) that matches
     * iText's line height for Noto Sans SC at the given half-point font size.
     *
     * Formula: round((1448 / 1000) × (szHalfPts / 2) × 20)
     */
    static int lineValForSz(int szHalfPts) {
        double pt = szHalfPts / 2.0;
        return (int) Math.round((HHEA_CELL / UPM) * pt * 20.0);
    }

    // -------------------------------------------------------------------------
    // I/O helpers
    // -------------------------------------------------------------------------

    private static byte[] readAllBytes(InputStream in) throws IOException {
        ByteArrayOutputStream buf = new ByteArrayOutputStream();
        byte[] chunk = new byte[8192];
        int n;
        while ((n = in.read(chunk)) != -1) buf.write(chunk, 0, n);
        return buf.toByteArray();
    }

    private static void copyStream(InputStream in, OutputStream out) throws IOException {
        byte[] buf = new byte[8192];
        int n;
        while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
    }

    private static void copyFile(File src, File dst) throws IOException {
        try (FileInputStream  fis = new FileInputStream(src);
             FileOutputStream fos = new FileOutputStream(dst)) {
            copyStream(fis, fos);
        }
    }

    // -------------------------------------------------------------------------
    // Optional: call directly for testing
    // -------------------------------------------------------------------------

    public static void main(String[] args) throws IOException {
        if (args.length < 2) {
            System.err.println("Usage: DocxLineSpacingFixer <input.docx> <output.docx>");
            System.exit(1);
        }
        fix(args[0], args[1]);
        System.out.println("Done: " + args[1]);
    }
}
