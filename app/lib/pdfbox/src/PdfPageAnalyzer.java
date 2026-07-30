import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;

import javax.imageio.stream.MemoryCacheImageOutputStream;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * V5 Stage 4: content-pdf をラスタライズして、シート単位の変更判定に使うハッシュを出す。
 *
 * 重要な約束:
 *  - ピクセルハッシュの入力形式を固定する。BufferedImage の内部形式をそのままハッシュすると、
 *    画像タイプや alpha premultiplication の違いで環境間・バージョン間に不安定さが出る。
 *      [幅(4byte BE)][高さ(4byte BE)][各画素を行優先で R,G,B の3バイト連結]  ← alpha は含めない
 *  - ファイル/ピクセルのハッシュは 64文字・大文字・prefix なし(server.ps1 の New-Sha256 と同形式)。
 *  - PNG は保存しない。ハッシュ計算後に画像を破棄する。
 *  - 解析の失敗は 1 シートの失敗にとどめ、プロセス全体を失敗させない。
 */
public final class PdfPageAnalyzer {

    public static final int ANALYZER_VERSION = 1;

    public static void main(String[] args) {
        String inputPath = null;
        String outputPath = null;
        int dpi = 150;
        for (int i = 0; i < args.length; i++) {
            if ("--input".equals(args[i]) && i + 1 < args.length) { inputPath = args[++i]; }
            else if ("--output".equals(args[i]) && i + 1 < args.length) { outputPath = args[++i]; }
            else if ("--dpi".equals(args[i]) && i + 1 < args.length) { dpi = Integer.parseInt(args[++i]); }
        }
        if (inputPath == null || outputPath == null) {
            System.err.println("usage: PdfPageAnalyzer --input <request.json> --output <result.json> [--dpi 150]");
            System.exit(2);
            return;
        }
        try {
            String requestText = new String(Files.readAllBytes(Paths.get(inputPath)), StandardCharsets.UTF_8);
            List<String[]> sheets = parseRequest(requestText);
            StringBuilder sb = new StringBuilder();
            sb.append("{\"schemaVersion\":1,\"analyzerVersion\":").append(ANALYZER_VERSION)
              .append(",\"dpi\":").append(dpi)
              .append(",\"colorMode\":\"RGB\"")
              .append(",\"javaVersion\":").append(quote(System.getProperty("java.version")))
              .append(",\"javaVendor\":").append(quote(System.getProperty("java.vendor")))
              .append(",\"sheets\":[");
            boolean first = true;
            for (String[] sheet : sheets) {
                if (!first) { sb.append(','); }
                first = false;
                sb.append(analyseOne(sheet[0], sheet[1], dpi));
            }
            sb.append("]}");
            try (Writer w = new OutputStreamWriter(Files.newOutputStream(Paths.get(outputPath)), StandardCharsets.UTF_8)) {
                w.write(sb.toString());
            }
            System.exit(0);
        } catch (Exception e) {
            System.err.println("PdfPageAnalyzer failed: " + e);
            System.exit(1);
        }
    }

    /** request.json は {"sheets":[{"sheetName":"1","pdf":"..."}]} だけを読む簡易パーサ。 */
    private static List<String[]> parseRequest(String json) {
        List<String[]> out = new ArrayList<String[]>();
        String key1 = "\"sheetName\"";
        String key2 = "\"pdf\"";
        int idx = 0;
        while (true) {
            int a = json.indexOf(key1, idx);
            if (a < 0) { break; }
            String name = readStringValue(json, a + key1.length());
            int b = json.indexOf(key2, a);
            if (b < 0) { break; }
            String pdf = readStringValue(json, b + key2.length());
            out.add(new String[] { name, pdf });
            idx = b + key2.length();
        }
        return out;
    }

    private static String readStringValue(String json, int from) {
        int i = from;
        while (i < json.length() && json.charAt(i) != '"') { i++; }
        i++;
        StringBuilder sb = new StringBuilder();
        while (i < json.length()) {
            char c = json.charAt(i);
            if (c == '\\' && i + 1 < json.length()) {
                char n = json.charAt(i + 1);
                if (n == 'u' && i + 5 < json.length()) {
                    sb.append((char) Integer.parseInt(json.substring(i + 2, i + 6), 16));
                    i += 6;
                    continue;
                }
                if (n == 'n') { sb.append('\n'); }
                else if (n == 't') { sb.append('\t'); }
                else { sb.append(n); }
                i += 2;
                continue;
            }
            if (c == '"') { break; }
            sb.append(c);
            i++;
        }
        return sb.toString();
    }

    private static String analyseOne(String sheetName, String pdfPath, int dpi) {
        PDDocument doc = null;
        try {
            File f = new File(pdfPath);
            if (!f.isFile()) {
                return errorSheet(sheetName, "content-pdf が見つかりません");
            }
            doc = PDDocument.load(f);
            int pageCount = doc.getNumberOfPages();
            PDFRenderer renderer = new PDFRenderer(doc);
            float scale = dpi / 72f;
            List<String> pageHashes = new ArrayList<String>();
            MessageDigest sheetDigest = MessageDigest.getInstance("SHA-256");
            sheetDigest.update(intBytes(pageCount));
            for (int i = 0; i < pageCount; i++) {
                BufferedImage img = renderer.renderImage(i, scale, ImageType.RGB);
                byte[] normalized = normalizedPixels(img);
                img.flush();
                MessageDigest pageDigest = MessageDigest.getInstance("SHA-256");
                String hex = toHex(pageDigest.digest(normalized));
                pageHashes.add(hex);
                sheetDigest.update(normalized, 0, Math.min(8, normalized.length));
                sheetDigest.update(hex.getBytes(StandardCharsets.US_ASCII));
            }
            String textHash = "";
            try {
                PDFTextStripper stripper = new PDFTextStripper();
                String text = stripper.getText(doc);
                MessageDigest td = MessageDigest.getInstance("SHA-256");
                textHash = toHex(td.digest(normalizeText(text).getBytes(StandardCharsets.UTF_8)));
            } catch (Exception ignore) {
                textHash = "";
            }
            StringBuilder sb = new StringBuilder();
            sb.append("{\"sheetName\":").append(quote(sheetName))
              .append(",\"status\":\"ok\",\"pageCount\":").append(pageCount)
              .append(",\"pageHashes\":[");
            for (int i = 0; i < pageHashes.size(); i++) {
                if (i > 0) { sb.append(','); }
                sb.append(quote(pageHashes.get(i)));
            }
            sb.append("],\"sheetVisualHash\":").append(quote(toHex(sheetDigest.digest())))
              .append(",\"textHash\":").append(quote(textHash))
              .append('}');
            return sb.toString();
        } catch (Throwable t) {
            // 1 シートの失敗でプロセス全体を落とさない。呼び出し側は status=error を unknown として扱う。
            return errorSheet(sheetName, String.valueOf(t));
        } finally {
            if (doc != null) { try { doc.close(); } catch (Exception ignore) { } }
        }
    }

    /** [幅(4)][高さ(4)][R,G,B * 画素数] 行優先。alpha は含めない。 */
    private static byte[] normalizedPixels(BufferedImage img) {
        int w = img.getWidth();
        int h = img.getHeight();
        byte[] out = new byte[8 + (w * h * 3)];
        byte[] wb = intBytes(w);
        byte[] hb = intBytes(h);
        System.arraycopy(wb, 0, out, 0, 4);
        System.arraycopy(hb, 0, out, 4, 4);
        int p = 8;
        int[] row = new int[w];
        for (int y = 0; y < h; y++) {
            img.getRGB(0, y, w, 1, row, 0, w);
            for (int x = 0; x < w; x++) {
                int rgb = row[x];
                out[p++] = (byte) ((rgb >> 16) & 0xFF);
                out[p++] = (byte) ((rgb >> 8) & 0xFF);
                out[p++] = (byte) (rgb & 0xFF);
            }
        }
        return out;
    }

    private static String normalizeText(String text) {
        if (text == null) { return ""; }
        return text.replaceAll("\\s+", " ").trim();
    }

    private static byte[] intBytes(int v) {
        return new byte[] { (byte) ((v >> 24) & 0xFF), (byte) ((v >> 16) & 0xFF), (byte) ((v >> 8) & 0xFF), (byte) (v & 0xFF) };
    }

    /** server.ps1 の New-Sha256 と同じ形式(64文字・大文字・prefix なし)。 */
    private static String toHex(byte[] bytes) {
        StringBuilder sb = new StringBuilder(bytes.length * 2);
        for (byte b : bytes) {
            String s = Integer.toHexString(b & 0xFF);
            if (s.length() == 1) { sb.append('0'); }
            sb.append(s);
        }
        return sb.toString().toUpperCase(java.util.Locale.ROOT);
    }

    private static String errorSheet(String sheetName, String message) {
        return "{\"sheetName\":" + quote(sheetName) + ",\"status\":\"error\",\"message\":" + quote(message) + "}";
    }

    private static String quote(String s) {
        if (s == null) { return "\"\""; }
        StringBuilder sb = new StringBuilder("\"");
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            switch (c) {
                case '"': sb.append("\\\""); break;
                case '\\': sb.append("\\\\"); break;
                case '\n': sb.append("\\n"); break;
                case '\r': sb.append("\\r"); break;
                case '\t': sb.append("\\t"); break;
                default:
                    if (c < 0x20 || c > 0x7E) {
                        sb.append(String.format("\\u%04x", (int) c));
                    } else {
                        sb.append(c);
                    }
            }
        }
        return sb.append('"').toString();
    }

    private PdfPageAnalyzer() { }
}
