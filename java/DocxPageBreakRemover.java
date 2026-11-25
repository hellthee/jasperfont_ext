package com.example.jasper;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * DocxPageBreakRemover
 *
 * Removes page breaks in a JasperReports .docx export that lead only to a
 * blank page. Page breaks before pages with content are kept.
 *
 * BACKGROUND:
 *   The JasperReports docx exporter (6.6.0 and 6.21.5 alike) writes each
 *   report page as a <w:tbl> and puts <w:pageBreakBefore/> in the first
 *   paragraph of every page after the first. Those breaks are real page
 *   boundaries and must normally stay: removing them makes Word run
 *   separate report pages together, so the document no longer matches the
 *   PDF.
 *
 *   A break is removed only when the page it starts is blank: no visible
 *   text (a <w:t> with a non-whitespace character) and no image
 *   (<w:drawing> or <w:pict>). Such pages show up in Word as trailing or
 *   interstitial blank pages.
 *
 *   A page extends from its break to the next break, or to the end of the
 *   body for the last page.
 *
 * USAGE:
 *   JasperExportManager.exportReportToDocxFile(jasperPrint, path);
 *   DocxPageBreakRemover.fixInPlace(path);
 *
 * REQUIREMENTS:
 *   Java 8+, no additional dependencies.
 */
public class DocxPageBreakRemover {

    private static final Pattern PAGE_BREAK = Pattern.compile("<w:pageBreakBefore/>");

    /** A text run with at least one visible character. Excludes w:tbl, w:tc. */
    private static final Pattern VISIBLE_TEXT = Pattern.compile(
            "<w:t(?:\\s[^>]*)?>[^<]*\\S[^<]*</w:t>");

    /**
     * Remove page breaks that lead only to blank pages.
     *
     * @param inputPath  JasperReports .docx export
     * @param outputPath where to write the result (may equal inputPath)
     */
    public static void fix(String inputPath, String outputPath) throws IOException {
        File outputFile = new File(outputPath).getAbsoluteFile();
        File tempFile   = File.createTempFile("docx_pbr_", ".tmp", outputFile.getParentFile());
        int removed = 0;
        int kept    = 0;

        try (ZipInputStream  zin  = new ZipInputStream(new FileInputStream(inputPath));
             ZipOutputStream zout = new ZipOutputStream(new FileOutputStream(tempFile))) {

            byte[] chunk = new byte[8192];
            ZipEntry entry;
            while ((entry = zin.getNextEntry()) != null) {
                ByteArrayOutputStream buf = new ByteArrayOutputStream();
                int n;
                while ((n = zin.read(chunk)) != -1) {
                    buf.write(chunk, 0, n);
                }
                byte[] data = buf.toByteArray();

                if ("word/document.xml".equals(entry.getName())) {
                    String xml = new String(data, StandardCharsets.UTF_8);

                    // pass 1: count breaks, so the arrays can be sized directly
                    Matcher m = PAGE_BREAK.matcher(xml);
                    int total = 0;
                    while (m.find()) {
                        total++;
                    }
                    int[] starts = new int[total];
                    int[] ends   = new int[total];
                    m.reset();
                    int i = 0;
                    while (m.find()) {
                        starts[i] = m.start();
                        ends[i]   = m.end();
                        i++;
                    }

                    int bodyEnd = xml.lastIndexOf("<w:sectPr");
                    if (bodyEnd < 0) {
                        bodyEnd = xml.lastIndexOf("</w:body>");
                    }
                    if (bodyEnd < 0) {
                        bodyEnd = xml.length();
                    }

                    // pass 2: drop a break only if its page has nothing visible
                    StringBuilder out = new StringBuilder(xml.length());
                    int pos = 0;
                    for (i = 0; i < total; i++) {
                        int pageEnd = (i + 1 < total) ? starts[i + 1] : bodyEnd;
                        String page = xml.substring(ends[i], Math.max(ends[i], pageEnd));
                        boolean blank = !VISIBLE_TEXT.matcher(page).find()
                                && page.indexOf("<w:drawing") < 0
                                && page.indexOf("<w:pict") < 0;
                        if (blank) {
                            out.append(xml, pos, starts[i]);
                            pos = ends[i];
                            removed++;
                        } else {
                            kept++;
                        }
                    }
                    out.append(xml, pos, xml.length());
                    data = out.toString().getBytes(StandardCharsets.UTF_8);
                }

                zout.putNextEntry(new ZipEntry(entry.getName()));
                zout.write(data);
                zout.closeEntry();
                zin.closeEntry();
            }
        } catch (IOException e) {
            tempFile.delete();
            throw e;
        }

        Files.move(tempFile.toPath(), outputFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
        System.out.printf("[DocxPageBreakRemover] removed %d break(s) before blank pages,"
                + " kept %d before pages with content%n", removed, kept);
    }

    /** Remove blank-page breaks in place. */
    public static void fixInPlace(String filePath) throws IOException {
        fix(filePath, filePath);
    }

    public static void main(String[] args) throws IOException {
        if (args.length < 1) {
            System.err.println("Usage: DocxPageBreakRemover <input.docx> [output.docx]");
            System.err.println("       If output is omitted, fixes in place.");
            System.exit(1);
        }
        fix(args[0], args.length > 1 ? args[1] : args[0]);
    }
}
