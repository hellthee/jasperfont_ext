package com.example.jasper;

import net.sf.jasperreports.engine.*;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import net.sf.jasperreports.engine.export.JRPdfExporter;
import net.sf.jasperreports.export.*;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;

/**
 * Example: Export a JasperReport with Chinese characters to PDF.
 *
 * Prerequisites:
 *   1. jasperreports-6.6.0.jar on classpath
 *   2. jasperreports-chinese-fonts-1.0.0.jar on classpath  <-- this extension
 *   3. itext-2.1.7.js6.jar (ships with JasperReports 6.6.0)
 *   4. commons-logging, commons-collections, commons-beanutils, etc.
 */
public class ChinesePdfExportExample {

    public static void main(String[] args) throws Exception {

        // ── 1. Verify the Chinese font is registered ──────────────────────
        verifyFontRegistration();

        // ── 2. Build a simple data source with Chinese content ────────────
        List<ProductRecord> records = buildSampleData();
        JRBeanCollectionDataSource dataSource = new JRBeanCollectionDataSource(records);

        // ── 3. Set report parameters ──────────────────────────────────────
        Map<String, Object> params = new HashMap<>();
        params.put("reportTitle", "销售报告 / Sales Report");
        params.put("companyName", "示例贸易有限公司 Example Trading Co., Ltd.");

        // ── 4. Compile + fill the report ──────────────────────────────────
        String jrxmlPath = "ChineseReportDemo.jrxml";
        JasperReport jasperReport = JasperCompileManager.compileReport(jrxmlPath);
        JasperPrint jasperPrint  = JasperFillManager.fillReport(jasperReport, params, dataSource);

        // ── 5. Export to PDF ──────────────────────────────────────────────
        String outputPath = "chinese-report-output.pdf";
        exportToPdf(jasperPrint, outputPath);

        System.out.println("PDF exported successfully: " + outputPath);
    }

    // -------------------------------------------------------------------------

    /**
     * Export JasperPrint to a PDF file.
     * The pdfEncoding="Identity-H" and pdfEmbedded="true" are already set
     * in the JRXML and fonts.xml, so no special exporter config is needed.
     */
    private static void exportToPdf(JasperPrint jasperPrint, String outputPath)
            throws JRException, FileNotFoundException {

        JRPdfExporter exporter = new JRPdfExporter();

        exporter.setExporterInput(
            new SimpleExporterInput(jasperPrint)
        );
        exporter.setExporterOutput(
            new SimpleOutputStreamExporterOutput(new FileOutputStream(outputPath))
        );

        // Optional PDF metadata
        SimplePdfExporterConfiguration config = new SimplePdfExporterConfiguration();
        config.setCreatorInfo("jasperreports-chinese-fonts-extension v1.0.0");
        config.setSubject("Chinese Character PDF Export Demo");
        exporter.setConfiguration(config);

        exporter.exportReport();
    }

    /**
     * Print all registered font families — useful for debugging.
     * You should see "Chinese" and "Noto Sans SC" in the output.
     */
    private static void verifyFontRegistration() {
        System.out.println("=== Registered JasperReports font families ===");
        FontUtil fontUtil = FontUtil.getInstance(DefaultJasperReportsContext.getInstance());
        for (net.sf.jasperreports.engine.fonts.FontFamily ff : fontUtil.getFontFamilies()) {
            System.out.println("  " + ff.getName());
        }
        System.out.println("==============================================");
    }

    // ─── Sample Data ─────────────────────────────────────────────────────────

    private static List<ProductRecord> buildSampleData() {
        List<ProductRecord> list = new ArrayList<>();
        list.add(new ProductRecord("Laptop",        "笔记本电脑",  5,  new BigDecimal("6999.00")));
        list.add(new ProductRecord("Keyboard",      "键盘",        10, new BigDecimal("299.00")));
        list.add(new ProductRecord("Monitor",       "显示器",      3,  new BigDecimal("1899.00")));
        list.add(new ProductRecord("Mouse",         "鼠标",        10, new BigDecimal("129.00")));
        list.add(new ProductRecord("USB-C Hub",     "集线器",      8,  new BigDecimal("199.00")));
        list.add(new ProductRecord("Webcam",        "摄像头",      4,  new BigDecimal("449.00")));
        list.add(new ProductRecord("Headphones",    "耳机",        6,  new BigDecimal("599.00")));
        list.add(new ProductRecord("Desk Lamp",     "台灯",        12, new BigDecimal("89.00")));
        return list;
    }

    // ─── Data Bean ───────────────────────────────────────────────────────────

    public static class ProductRecord {
        private String     productName;
        private String     productNameCN;
        private Integer    quantity;
        private BigDecimal unitPrice;
        private BigDecimal total;

        public ProductRecord(String en, String cn, int qty, BigDecimal price) {
            this.productName   = en;
            this.productNameCN = cn;
            this.quantity      = qty;
            this.unitPrice     = price;
            this.total         = price.multiply(new BigDecimal(qty));
        }

        // Getters — required by JRBeanCollectionDataSource
        public String     getProductName()   { return productName; }
        public String     getProductNameCN() { return productNameCN; }
        public Integer    getQuantity()      { return quantity; }
        public BigDecimal getUnitPrice()     { return unitPrice; }
        public BigDecimal getTotal()         { return total; }
    }
}
