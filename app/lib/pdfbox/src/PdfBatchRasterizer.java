import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

/**
 * Rasterizes every before/after PDF for one diff job in a single JVM.
 * Independent PDFs are rendered concurrently with a bounded worker pool.
 */
public final class PdfBatchRasterizer {
    public static void main(String[] args) {
        String requestPath = null;
        String outputPath = null;
        int dpi = 120;
        int threads = Math.max(1, Math.min(4, Runtime.getRuntime().availableProcessors()));
        for (int i = 0; i < args.length; i++) {
            if ("--request".equals(args[i]) && i + 1 < args.length) requestPath = args[++i];
            else if ("--output".equals(args[i]) && i + 1 < args.length) outputPath = args[++i];
            else if ("--dpi".equals(args[i]) && i + 1 < args.length) dpi = Integer.parseInt(args[++i]);
            else if ("--threads".equals(args[i]) && i + 1 < args.length) threads = Integer.parseInt(args[++i]);
        }
        if (requestPath == null || outputPath == null) {
            System.err.println("usage: PdfBatchRasterizer --request <tsv> --output <dir> [--dpi 120] [--threads 4]");
            System.exit(2);
            return;
        }
        try {
            final int renderDpi = dpi;
            final Path root = Paths.get(outputPath);
            Files.createDirectories(root);
            List<String> lines = Files.readAllLines(Paths.get(requestPath), StandardCharsets.UTF_8);
            ExecutorService pool = Executors.newFixedThreadPool(Math.max(1, Math.min(4, threads)));
            List<Future<?>> futures = new ArrayList<Future<?>>();
            try {
                for (String line : lines) {
                    if (line.trim().isEmpty()) continue;
                    String[] fields = line.split("\\t", -1);
                    if (fields.length != 3 || !fields[0].matches("[A-Za-z0-9_-]+")) {
                        throw new IllegalArgumentException("invalid raster request line");
                    }
                    final String id = fields[0];
                    final String before = decode(fields[1]);
                    final String after = decode(fields[2]);
                    if (!before.isEmpty()) futures.add(pool.submit(() -> renderSafely(root, id, "before", before, renderDpi)));
                    if (!after.isEmpty()) futures.add(pool.submit(() -> renderSafely(root, id, "after", after, renderDpi)));
                }
                for (Future<?> future : futures) future.get();
            } finally {
                pool.shutdownNow();
            }
            System.exit(0);
        } catch (Throwable t) {
            System.err.println("PdfBatchRasterizer failed: " + t);
            System.exit(1);
        }
    }

    private static String decode(String encoded) {
        if (encoded == null || encoded.isEmpty()) return "";
        return new String(Base64.getDecoder().decode(encoded), StandardCharsets.UTF_8);
    }

    private static void renderSafely(Path root, String id, String side, String pdfPath, int dpi) {
        Path dir = root.resolve(id);
        try {
            Files.createDirectories(dir);
            renderPdf(new File(pdfPath), dir, side, dpi);
        } catch (Throwable t) {
            try {
                Files.write(dir.resolve(side + ".error.txt"), String.valueOf(t).getBytes(StandardCharsets.UTF_8));
            } catch (Exception ignored) { }
        }
    }

    private static void renderPdf(File pdf, Path output, String prefix, int dpi) throws Exception {
        if (!pdf.isFile()) throw new IllegalArgumentException("PDF not found: " + pdf);
        try (PDDocument document = PDDocument.load(pdf)) {
            PDFRenderer renderer = new PDFRenderer(document);
            float scale = dpi / 72f;
            for (int i = 0; i < document.getNumberOfPages(); i++) {
                BufferedImage image = renderer.renderImage(i, scale, ImageType.RGB);
                try {
                    File target = output.resolve(prefix + "-" + (i + 1) + ".png").toFile();
                    if (!ImageIO.write(image, "png", target)) throw new IllegalStateException("PNG writer unavailable");
                } finally {
                    image.flush();
                }
            }
        }
    }

    private PdfBatchRasterizer() { }
}
