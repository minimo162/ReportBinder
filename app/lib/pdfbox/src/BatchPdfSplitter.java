import java.io.*;
import java.lang.reflect.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.*;

/**
 * BatchPdfSplitter
 *
 * Excel can export multiple selected worksheets to one PDF much faster than exporting
 * each sheet separately. This helper splits that combined PDF back into one content PDF per sheet.
 * It uses reflection so it can be compiled without bundling Apache PDFBox; at runtime the
 * classpath must include pdfbox-app.jar, same as ReportPdfComposer.
 *
 * Usage:
 *   java -cp "ReportPdfComposer.jar;pdfbox-app.jar" BatchPdfSplitter --source combined.pdf --map map.tsv
 *
 * map.tsv: outputPdf<TAB>startPage1Based<TAB>pageCount
 */
public class BatchPdfSplitter {
    public static void main(String[] args) throws Exception {
        if (args.length != 4 || !"--source".equals(args[0]) || !"--map".equals(args[2])) {
            System.err.println("Usage: BatchPdfSplitter --source combined.pdf --map map.tsv");
            System.exit(2);
        }
        long startedAt = System.currentTimeMillis();
        File source = new File(args[1]);
        File map = new File(args[3]);
        if (!source.isFile()) throw new FileNotFoundException(source.getAbsolutePath());
        if (!map.isFile()) throw new FileNotFoundException(map.getAbsolutePath());

        PdfBox box = new PdfBox();
        Object src = box.load(source);
        int written = 0;
        try {
            int totalPages = box.getNumberOfPages(src);
            List<String> lines = Files.readAllLines(map.toPath(), StandardCharsets.UTF_8);
            for (String line : lines) {
                if (line == null) continue;
                line = line.trim();
                if (line.length() == 0 || line.startsWith("#")) continue;
                String[] parts = line.split("\\t", -1);
                if (parts.length < 3) throw new IllegalArgumentException("Bad map line: " + line);
                File out = new File(parts[0]);
                int start = Integer.parseInt(parts[1]);
                int count = Integer.parseInt(parts[2]);
                if (start < 1 || count < 1 || start + count - 1 > totalPages) {
                    throw new IllegalArgumentException("Page range is outside source PDF: " + line + " total=" + totalPages);
                }
                File parent = out.getParentFile();
                if (parent != null) parent.mkdirs();
                Object dst = box.newDocument();
                try {
                    for (int i = 0; i < count; i++) {
                        Object page = box.getPage(src, start - 1 + i);
                        box.importPage(dst, page);
                    }
                    box.save(dst, out);
                    written++;
                } finally {
                    box.close(dst);
                }
            }
        } finally {
            box.close(src);
        }
        long elapsedMs = System.currentTimeMillis() - startedAt;
        System.out.println("split: " + source.getAbsolutePath() + " files=" + written + " elapsedMs=" + elapsedMs);
    }

    static class PdfBox {
        private final Class<?> pdDocumentClass;
        private final Class<?> pdPageClass;
        private final Method loadFile;
        private final Method loadString;
        private final Method loaderLoadFile;
        private final Constructor<?> ctor;
        private final Method getNumberOfPages;
        private final Method getPage;
        private final Method importPage;
        private final Method saveFile;
        private final Method saveString;
        private final Method close;

        PdfBox() throws Exception {
            pdDocumentClass = Class.forName("org.apache.pdfbox.pdmodel.PDDocument");
            pdPageClass = Class.forName("org.apache.pdfbox.pdmodel.PDPage");
            ctor = pdDocumentClass.getConstructor();
            loadFile = methodOrNull(pdDocumentClass, "load", File.class);
            loadString = methodOrNull(pdDocumentClass, "load", String.class);
            Method loader = null;
            try {
                Class<?> loaderClass = Class.forName("org.apache.pdfbox.Loader");
                loader = methodOrNull(loaderClass, "loadPDF", File.class);
            } catch (Throwable ignored) { }
            loaderLoadFile = loader;
            getNumberOfPages = pdDocumentClass.getMethod("getNumberOfPages");
            getPage = pdDocumentClass.getMethod("getPage", Integer.TYPE);
            importPage = pdDocumentClass.getMethod("importPage", pdPageClass);
            saveFile = methodOrNull(pdDocumentClass, "save", File.class);
            saveString = methodOrNull(pdDocumentClass, "save", String.class);
            close = pdDocumentClass.getMethod("close");
        }

        Object load(File f) throws Exception {
            if (loadFile != null) return loadFile.invoke(null, f);
            if (loaderLoadFile != null) return loaderLoadFile.invoke(null, f);
            if (loadString != null) return loadString.invoke(null, f.getAbsolutePath());
            throw new NoSuchMethodException("No compatible PDFBox load method found");
        }
        Object newDocument() throws Exception { return ctor.newInstance(); }
        int getNumberOfPages(Object doc) throws Exception { return ((Number)getNumberOfPages.invoke(doc)).intValue(); }
        Object getPage(Object doc, int index) throws Exception { return getPage.invoke(doc, index); }
        void importPage(Object doc, Object page) throws Exception { importPage.invoke(doc, page); }
        void save(Object doc, File f) throws Exception {
            if (saveFile != null) { saveFile.invoke(doc, f); return; }
            if (saveString != null) { saveString.invoke(doc, f.getAbsolutePath()); return; }
            throw new NoSuchMethodException("No compatible PDFBox save method found");
        }
        void close(Object doc) {
            if (doc == null) return;
            try { close.invoke(doc); } catch (Throwable ignored) { }
        }
        static Method methodOrNull(Class<?> c, String name, Class<?>... args) {
            try { return c.getMethod(name, args); } catch (Throwable t) { return null; }
        }
    }
}
