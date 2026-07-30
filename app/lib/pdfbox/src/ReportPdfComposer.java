import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.*;

import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.util.Matrix;

/**
 * ReportPdfComposer
 *
 * PDFBox 2.x を前提にした専用CLIです。
 * 既存 content-pdf を結合し、必要に応じて本文を左右へ移動し、ページ番号を物理ページ下部中央へ描画します。
 *
 * Usage:
 *   java -cp "ReportPdfComposer.jar;pdfbox-app.jar" ReportPdfComposer --manifest manifest.json
 */
public class ReportPdfComposer {
    @SuppressWarnings("unchecked")
    public static void main(String[] args) throws Exception {
        if (args.length != 2 || !"--manifest".equals(args[0])) {
            System.err.println("Usage: java -cp \"ReportPdfComposer.jar;pdfbox-app.jar\" ReportPdfComposer --manifest manifest.json");
            System.exit(2);
        }
        long startedAt = System.currentTimeMillis();
        File manifestFile = new File(args[1]);
        Map<String, Object> manifest = (Map<String, Object>) MiniJson.parse(readAll(manifestFile));
        String outputPdf = str(manifest.get("outputPdf"));
        if (outputPdf.isEmpty()) throw new IllegalArgumentException("outputPdf is required");
        List<Object> pages = (List<Object>) manifest.get("pages");
        if (pages == null || pages.isEmpty()) throw new IllegalArgumentException("pages is empty");

        Map<String, Object> pageNumber = (Map<String, Object>) manifest.get("pageNumber");
        float fontSize = pageNumber == null ? 8f : flt(pageNumber.get("fontSize"), 8f);
        float bottomPt = pageNumber == null ? 18f : flt(pageNumber.get("bottomPt"), 18f);

        File outFile = new File(outputPdf);
        File parent = outFile.getParentFile();
        if (parent != null) parent.mkdirs();

        int physicalPageNo = 0;
        try (PDDocument out = new PDDocument()) {
            PDFont pageFont = loadArialOrFallback(out);
            PDFMergerUtility merger = new PDFMergerUtility();

            for (Object o : pages) {
                Map<String, Object> entry = (Map<String, Object>) o;
                File source = new File(str(entry.get("sourcePdf")));
                String numberingMode = str(entry.get("numberingMode"));
                if (numberingMode.isEmpty()) numberingMode = "visible";
                if (!"visible".equals(numberingMode) && !"none".equals(numberingMode)) {
                    throw new IllegalArgumentException("Unsupported numberingMode: " + numberingMode);
                }
                float punchShiftPt = flt(entry.get("punchShiftPt"), 0f);
                if (!source.isFile()) throw new FileNotFoundException(source.getAbsolutePath());

                try (PDDocument src = PDDocument.load(source)) {
                    int before = out.getNumberOfPages();
                    int sourcePageCount = src.getNumberOfPages();
                    merger.appendDocument(out, src);

                    for (int i = 0; i < sourcePageCount; i++) {
                        physicalPageNo++;
                        PDPage newPage = out.getPage(before + i);

                        float shift = 0f;
                        if (punchShiftPt != 0f) shift = (physicalPageNo % 2 == 1) ? punchShiftPt : -punchShiftPt;
                        if (shift != 0f) wrapExistingPageContent(out, newPage, shift);

                        if ("visible".equals(numberingMode)) {
                            drawPageNumber(out, newPage, pageFont, fontSize, bottomPt, formatPageNumber(physicalPageNo, pageNumber));
                        }
                    }
                }
            }
            out.save(outFile);
        }
        long elapsedMs = System.currentTimeMillis() - startedAt;
        System.out.println("created: " + outFile.getAbsolutePath() + " pages=" + physicalPageNo + " elapsedMs=" + elapsedMs + " mode=merge-wrap");
    }

    private static void wrapExistingPageContent(PDDocument doc, PDPage page, float shift) throws IOException {
        try (PDPageContentStream pre = new PDPageContentStream(doc, page, PDPageContentStream.AppendMode.PREPEND, true, false)) {
            pre.saveGraphicsState();
            pre.transform(Matrix.getTranslateInstance(shift, 0));
        }
        try (PDPageContentStream post = new PDPageContentStream(doc, page, PDPageContentStream.AppendMode.APPEND, true, false)) {
            post.restoreGraphicsState();
        }
    }

    private static String formatPageNumber(int pageNo, Map<String, Object> config) {
        String format = config == null ? "hyphenated" : str(config.get("format"));
        if (format == null || format.length() == 0 || "hyphenated".equals(format)) return "- " + pageNo + " -";
        if ("plain".equals(format)) return String.valueOf(pageNo);
        return "- " + pageNo + " -";
    }

    private static void drawPageNumber(PDDocument doc, PDPage page, PDFont font, float size, float bottomPt, String text) throws IOException {
        PDRectangle box = page.getMediaBox();
        float width = font.getStringWidth(text) / 1000f * size;
        float x = box.getLowerLeftX() + (box.getWidth() - width) / 2f;
        float y = box.getLowerLeftY() + bottomPt;
        try (PDPageContentStream cs = new PDPageContentStream(doc, page, PDPageContentStream.AppendMode.APPEND, true, true)) {
            cs.beginText();
            cs.setFont(font, size);
            cs.newLineAtOffset(x, y);
            cs.showText(text);
            cs.endText();
        }
    }

    private static PDFont loadArialOrFallback(PDDocument doc) throws IOException {
        String windir = System.getenv("WINDIR");
        List<File> candidates = new ArrayList<>();
        if (windir != null && !windir.isEmpty()) candidates.add(new File(windir, "Fonts/arial.ttf"));
        candidates.add(new File("C:/Windows/Fonts/arial.ttf"));
        candidates.add(new File("/usr/share/fonts/truetype/msttcorefonts/Arial.ttf"));
        for (File f : candidates) {
            if (f.isFile()) {
                PDFont font = tryLoadType0Font(doc, f);
                if (font != null) return font;
            }
        }
        System.err.println("WARN: Arial font was not available or could not be loaded by this PDFBox version. Falling back to Helvetica.");
        return PDType1Font.HELVETICA;
    }

    /**
     * Load Arial without binding the class file to a single PDFBox overload.
     * Some PDFBox 2.x app jars do not have PDType0Font.load(PDDocument, File, boolean).
     * Calling it directly causes NoSuchMethodError at runtime, so we probe compatible
     * overloads by reflection and fall back safely when none is available.
     */
    private static PDFont tryLoadType0Font(PDDocument doc, File fontFile) {
        try {
            Class<?> cls = Class.forName("org.apache.pdfbox.pdmodel.font.PDType0Font");
            try {
                java.lang.reflect.Method m = cls.getMethod("load", PDDocument.class, File.class);
                Object loaded = m.invoke(null, doc, fontFile);
                if (loaded instanceof PDFont) return (PDFont) loaded;
            } catch (NoSuchMethodException ignored) {
                // Try the stream based overload below.
            }
            try {
                java.lang.reflect.Method m = cls.getMethod("load", PDDocument.class, InputStream.class, Boolean.TYPE);
                FileInputStream in = new FileInputStream(fontFile);
                try {
                    Object loaded = m.invoke(null, doc, in, Boolean.TRUE);
                    if (loaded instanceof PDFont) return (PDFont) loaded;
                } finally {
                    in.close();
                }
            } catch (NoSuchMethodException ignored) {
                // Try the older stream based overload below.
            }
            try {
                java.lang.reflect.Method m = cls.getMethod("load", PDDocument.class, InputStream.class);
                FileInputStream in = new FileInputStream(fontFile);
                try {
                    Object loaded = m.invoke(null, doc, in);
                    if (loaded instanceof PDFont) return (PDFont) loaded;
                } finally {
                    in.close();
                }
            } catch (NoSuchMethodException ignored) {
                // Fall back to Helvetica.
            }
        } catch (Throwable t) {
            String msg = t.getMessage();
            if (msg == null || msg.length() == 0) msg = t.getClass().getName();
            System.err.println("WARN: Could not load Arial with PDFBox: " + msg);
        }
        return null;
    }

    private static String readAll(File f) throws IOException {
        String text = new String(Files.readAllBytes(f.toPath()), StandardCharsets.UTF_8);
        if (!text.isEmpty() && text.charAt(0) == '\uFEFF') text = text.substring(1);
        return text;
    }
    private static String str(Object o) { return o == null ? "" : String.valueOf(o); }
    private static float flt(Object o, float def) {
        if (o == null) return def;
        if (o instanceof Number) return ((Number)o).floatValue();
        try { return Float.parseFloat(String.valueOf(o)); } catch (Exception e) { return def; }
    }

    /** Minimal JSON parser for this manifest format. */
    static class MiniJson {
        private final String s;
        private int p = 0;
        private MiniJson(String s) { this.s = s; }
        static Object parse(String s) { return new MiniJson(s).parseValue(); }
        private Object parseValue() {
            ws();
            if (p >= s.length()) throw err("unexpected end");
            char c = s.charAt(p);
            if (c == '{') return obj();
            if (c == '[') return arr();
            if (c == '"') return string();
            if (c == 't' || c == 'f') return bool();
            if (c == 'n') { expect("null"); return null; }
            return num();
        }
        private Map<String,Object> obj() {
            Map<String,Object> m = new LinkedHashMap<>();
            p++; ws();
            if (peek('}')) { p++; return m; }
            while (true) {
                ws(); String k = string(); ws(); ch(':'); Object v = parseValue(); m.put(k, v); ws();
                if (peek('}')) { p++; return m; }
                ch(',');
            }
        }
        private List<Object> arr() {
            List<Object> a = new ArrayList<>();
            p++; ws();
            if (peek(']')) { p++; return a; }
            while (true) {
                a.add(parseValue()); ws();
                if (peek(']')) { p++; return a; }
                ch(',');
            }
        }
        private String string() {
            ch('"');
            StringBuilder b = new StringBuilder();
            while (p < s.length()) {
                char c = s.charAt(p++);
                if (c == '"') return b.toString();
                if (c == '\\') {
                    if (p >= s.length()) throw err("bad escape");
                    char e = s.charAt(p++);
                    switch (e) {
                        case '"': b.append('"'); break;
                        case '\\': b.append('\\'); break;
                        case '/': b.append('/'); break;
                        case 'b': b.append('\b'); break;
                        case 'f': b.append('\f'); break;
                        case 'n': b.append('\n'); break;
                        case 'r': b.append('\r'); break;
                        case 't': b.append('\t'); break;
                        case 'u':
                            if (p + 4 > s.length()) throw err("bad unicode escape");
                            b.append((char)Integer.parseInt(s.substring(p, p+4), 16)); p += 4; break;
                        default: throw err("bad escape: " + e);
                    }
                } else b.append(c);
            }
            throw err("unterminated string");
        }
        private Boolean bool() {
            if (s.startsWith("true", p)) { p += 4; return Boolean.TRUE; }
            if (s.startsWith("false", p)) { p += 5; return Boolean.FALSE; }
            throw err("bad boolean");
        }
        private Number num() {
            int start = p;
            while (p < s.length() && "-+0123456789.eE".indexOf(s.charAt(p)) >= 0) p++;
            if (start == p) {
                char c = p < s.length() ? s.charAt(p) : '\0';
                throw err("unexpected token '" + c + "'");
            }
            String n = s.substring(start, p);
            if (n.indexOf('.') >= 0 || n.indexOf('e') >= 0 || n.indexOf('E') >= 0) return Double.parseDouble(n);
            try { return Integer.parseInt(n); } catch (NumberFormatException e) { return Long.parseLong(n); }
        }
        private void ws() { while (p < s.length() && (Character.isWhitespace(s.charAt(p)) || s.charAt(p) == '\uFEFF')) p++; }
        private boolean peek(char c) { return p < s.length() && s.charAt(p) == c; }
        private void ch(char c) { ws(); if (!peek(c)) throw err("expected '" + c + "'"); p++; }
        private void expect(String x) { if (!s.startsWith(x, p)) throw err("expected " + x); p += x.length(); }
        private RuntimeException err(String m) { return new IllegalArgumentException(m + " at " + p); }
    }
}
