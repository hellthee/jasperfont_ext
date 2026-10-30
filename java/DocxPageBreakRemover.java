package com.example.jasper;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.regex.*;
import java.util.zip.*;

/**
 * DocxPageBreakRemover
 *
 * Removes forced page breaks inserted by the JasperReports 6.6.0 docx exporter.
 *
 * PROBLEM:
 *   JasperReports exports each PDF page as a separate <w:tbl> (Word table).
 *   The first cell of every table after the first one contains:
 *     <w:pageBreakBefore/>
 *   inside its <w:pPr>, which forces Word to start a new page regardless
 *   of how much space remains — even when the entire report fits on one page.
 *
 * FIX:
 *   Strip all <w:pageBreakBefore/> elements from word/document.xml.
 *   Image paragraphs and empty cells are unaffected.
 *
 * USAGE — standalone after JasperReports export:
 *   DocxPageBreakRemover.fix("/tmp/report.docx", "/tmp/report_fixed.docx");
 *
 * USAGE — in-place:
 *   DocxPageBreakRemover.fixInPlace("/tmp/report.docx");
 *
 * USAGE — chain with DocxLineSpacingFixer:
 *   JasperExportManager.exportReportToDocxFile(jasperPrint, "/tmp/report.docx");
 *   DocxPageBreakRemover.fixInPlace("/tmp/report.docx");
 *   DocxLineSpacingFixer.fixInPlace("/tmp/report.docx");
 *
 * REQUIREMENTS:
 *   Java 8+, no additional dependencies.
 */
public class DocxPageBreakRemover {

    // Matches <w:pageBreakBefore/> with any surrounding whitespace/newlines
    private static final Pattern PAGE_BREAK_PATTERN = Pattern.compile(
            "\\s*<w:pageBreakBefore/>\\s*", Pattern.DOTALL);

    /**
     * Remove page breaks from a JasperReports .docx export.
     *
     * @param inputPath  path to the JR-exported .docx
     * @param outputPath path for the fixed output .docx
     * @throws IOException on read/write error
     */
    public static void fix(String inputPath, String outputPath) throws IOException {
        File inputFile  = new File(inputPath);
        File outputFile = new File(outputPath);
        File tempFile   = File.createTempFile("docx_pbr_", ".docx",
                outputFile.getParentFile());

        try (ZipInputStream  zin  = new ZipInputStream(new FileInputStream(inputFile));
             ZipOutputStream zout = new ZipOutputStream(new FileOutputStream(tempFile))) {

            ZipEntry entry;
            while ((entry = zin.getNextEntry()) != null) {
                zout.putNextEntry(new ZipEntry(entry.getName()));

                if ("word/document.xml".equals(entry.getName())) {
                    byte[] xmlBytes = readAllBytes(zin);
                    String xml      = new String(xmlBytes, StandardCharsets.UTF_8);
                    String fixed    = removePageBreaks(xml);
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
     * Remove page breaks in-place (overwrites the original file).
     *
     * @param filePath path to the .docx file to fix
     * @throws IOException on read/write error
     */
    public static void fixInPlace(String filePath) throws IOException {
        String tempPath = filePath + ".tmp";
        fix(filePath, tempPath);
        new File(filePath).delete();
        new File(tempPath).renameTo(new File(filePath));
    }

    // -------------------------------------------------------------------------
    // Core processing
    // -------------------------------------------------------------------------

    private static String removePageBreaks(String xml) {
        Matcher m = PAGE_BREAK_PATTERN.matcher(xml);
        int count = 0;
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            m.appendReplacement(sb, "");
            count++;
        }
        m.appendTail(sb);
        System.out.printf("[DocxPageBreakRemover] removed %d <w:pageBreakBefore/> element(s)%n",
                count);
        return sb.toString();
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
    // Command-line entry point
    // -------------------------------------------------------------------------

    public static void main(String[] args) throws IOException {
        if (args.length < 1) {
            System.err.println("Usage: DocxPageBreakRemover <input.docx> [output.docx]");
            System.err.println("       If output is omitted, fixes in-place.");
            System.exit(1);
        }
        if (args.length == 1) {
            fixInPlace(args[0]);
            System.out.println("Done (in-place): " + args[0]);
        } else {
            fix(args[0], args[1]);
            System.out.println("Done: " + args[1]);
        }
    }
}
