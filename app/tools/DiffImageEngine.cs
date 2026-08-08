using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

public sealed class ReportBinderDiffRegion
{
    public string regionId;
    public string kind;
    public double x;
    public double y;
    public double width;
    public double height;
    public double confidence;
    public int pixelCount;
}

public sealed class ReportBinderDiffPage
{
    public int pageNumber;
    public int width;
    public int height;
    public bool pageSizeChanged;
    public string status;
    public string message;
    public double confidence;
    public double changedRatio;
    public int regionCount;
    public ReportBinderDiffRegion[] regions;
    public string beforeFile;
    public string afterFile;
    public string beforeMaskFile;
    public string beforeOverlayFile;
    public string afterMaskFile;
    public string afterOverlayFile;
}

public sealed class ReportBinderDiffBatchPageRequest
{
    public string itemId;
    public string beforePath;
    public string afterPath;
    public string outputDirectory;
    public int pageNumber;
    public string kind;
    public bool copyBefore;
    public bool copyAfter;
}

public sealed class ReportBinderDiffBatchPageResult
{
    public string itemId;
    public int pageNumber;
    public ReportBinderDiffPage page;
    public string error;
}

public static class ReportBinderDiffEngine
{
    private const int MaximumRegionsPerPage = 120;
    private const int MaximumModifiedLabelsPerPage = 12;
    private const int MaximumAlignmentOffset = 4;

    public static double MeanAbsoluteByteDistance(byte[] before, byte[] after)
    {
        if (before == null || after == null || before.Length == 0 || before.Length != after.Length) return 1.0;
        long difference = 0;
        for (int i = 0; i < before.Length; i++) difference += Math.Abs((int)before[i] - (int)after[i]);
        return Math.Min(1.0, difference / (before.Length * 255.0));
    }

    private sealed class PixelRegion
    {
        public int minX;
        public int minY;
        public int maxX;
        public int maxY;
        public int count;
        public int added;
        public int removed;
        public int modified;
        public long diffTotal;
        public string kind;
        public double confidence;
    }

    private static Bitmap LoadImage(string path)
    {
        if (String.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;
        using (Image src = Image.FromFile(path))
        {
            return new Bitmap(src);
        }
    }

    private static Bitmap Normalize(Bitmap source, int width, int height)
    {
        Bitmap normalized = new Bitmap(width, height, PixelFormat.Format24bppRgb);
        normalized.SetResolution(150, 150);
        using (Graphics g = Graphics.FromImage(normalized))
        {
            g.Clear(Color.White);
            // DrawImageUnscaled scales by the source DPI metadata despite its name.
            // An explicit destination rectangle keeps this 1:1 in pixels even when
            // PDFBox omits or changes the PNG resolution metadata.
            if (source != null) g.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height));
        }
        return normalized;
    }

    private static byte[] ReadBgr(Bitmap bitmap)
    {
        Rectangle rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
        BitmapData data = bitmap.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
        try
        {
            int stride = Math.Abs(data.Stride);
            byte[] raw = new byte[stride * bitmap.Height];
            Marshal.Copy(data.Scan0, raw, 0, raw.Length);
            if (stride == bitmap.Width * 3) return raw;
            byte[] packed = new byte[bitmap.Width * bitmap.Height * 3];
            for (int y = 0; y < bitmap.Height; y++)
                Buffer.BlockCopy(raw, y * stride, packed, y * bitmap.Width * 3, bitmap.Width * 3);
            return packed;
        }
        finally
        {
            bitmap.UnlockBits(data);
        }
    }

    private static bool IsWhite(byte b, byte g, byte r)
    {
        return b >= 246 && g >= 246 && r >= 246;
    }

    private static Color RegionColor(string kind)
    {
        if (kind == "added") return Color.FromArgb(28, 103, 190);
        if (kind == "removed") return Color.FromArgb(200, 48, 55);
        if (kind == "unknown") return Color.FromArgb(105, 113, 124);
        return Color.FromArgb(212, 132, 0);
    }

    private static string RegionLabel(string kind)
    {
        if (kind == "added") return "A";
        if (kind == "removed") return "D";
        if (kind == "unknown") return "?";
        return "M";
    }

    private static bool VisibleOnSide(string kind, bool before)
    {
        if (kind == "modified" || kind == "unknown") return true;
        if (kind == "added") return !before;
        if (kind == "removed") return before;
        return true;
    }

    private static void SaveLayerImages(
        string outputDirectory,
        int pageNumber,
        int width,
        int height,
        IList<PixelRegion> regions,
        bool before,
        out string maskName,
        out string overlayName)
    {
        // Short names on purpose: the cache lives under
        // input-history\<workbookId>\<snapshotId>\renders\<versionId>\comparisons\d<key>\p\<sheetKey>\
        // and Windows PowerShell 5.1 still enforces MAX_PATH on the deep submission folders.
        string prefix = pageNumber.ToString("0000") + (before ? "-b" : "-a");
        maskName = prefix + "m.png";
        overlayName = prefix + "o.png";
        string maskPath = Path.Combine(outputDirectory, maskName);
        string overlayPath = Path.Combine(outputDirectory, overlayName);

        if (regions.Count == 0)
        {
            using (Bitmap empty = new Bitmap(1, 1, PixelFormat.Format32bppArgb))
            {
                empty.SetPixel(0, 0, Color.Transparent);
                empty.Save(maskPath, ImageFormat.Png);
                empty.Save(overlayPath, ImageFormat.Png);
            }
            return;
        }

        using (Bitmap mask = new Bitmap(width, height, PixelFormat.Format32bppArgb))
        using (Bitmap overlay = new Bitmap(width, height, PixelFormat.Format32bppArgb))
        using (Graphics mg = Graphics.FromImage(mask))
        using (Graphics og = Graphics.FromImage(overlay))
        {
            mg.Clear(Color.Transparent);
            og.Clear(Color.Transparent);
            og.SmoothingMode = SmoothingMode.AntiAlias;
            og.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

            foreach (PixelRegion region in regions)
            {
                if (!VisibleOnSide(region.kind, before)) continue;
                Color color = RegionColor(region.kind);
                Rectangle rect = Rectangle.FromLTRB(region.minX, region.minY, region.maxX + 1, region.maxY + 1);
                using (Brush fill = new SolidBrush(Color.FromArgb(255, color)))
                {
                    mg.FillRectangle(fill, rect);
                }
                using (Pen pen = new Pen(color, Math.Max(3f, width / 520f)))
                {
                    if (region.kind == "removed") pen.DashStyle = DashStyle.Dash;
                    else if (region.kind == "unknown") pen.DashStyle = DashStyle.Dot;
                    og.DrawRectangle(pen, rect);
                }

                bool drawLabel = region.kind != "modified" || regions.Count <= MaximumModifiedLabelsPerPage;
                if (!drawLabel) continue;
                string label = RegionLabel(region.kind);
                float labelSize = Math.Max(13f, Math.Min(22f, width / 70f));
                RectangleF labelBox = new RectangleF(rect.Left, Math.Max(0, rect.Top - labelSize), labelSize, labelSize);
                if (rect.Top < labelSize) labelBox.Y = rect.Top;
                using (Brush labelFill = new SolidBrush(color))
                using (Font font = new Font("Arial", Math.Max(10f, labelSize * 0.52f), FontStyle.Bold, GraphicsUnit.Pixel))
                using (StringFormat format = new StringFormat())
                {
                    format.Alignment = StringAlignment.Center;
                    format.LineAlignment = StringAlignment.Center;
                    og.FillRectangle(labelFill, labelBox);
                    og.DrawString(label, font, Brushes.White, labelBox, format);
                }
            }
            mask.Save(maskPath, ImageFormat.Png);
            overlay.Save(overlayPath, ImageFormat.Png);
        }
    }

    private static PixelRegion FullPageRegion(int width, int height, string kind)
    {
        int inset = Math.Max(6, Math.Min(width, height) / 250);
        return new PixelRegion
        {
            minX = inset,
            minY = inset,
            maxX = Math.Max(inset, width - inset - 1),
            maxY = Math.Max(inset, height - inset - 1),
            count = width * height,
            kind = kind,
            confidence = 1.0
        };
    }

    private static bool[] Dilate(bool[] source, int width, int height, int radius)
    {
        // Separable Chebyshev dilation (two sliding-window passes per axis).
        // Used only to decide which change pixels belong to the same region, so that
        // the digits of one changed number do not become one box per digit.
        if (radius <= 0) return source;
        bool[] horizontal = new bool[source.Length];
        for (int y = 0; y < height; y++)
        {
            int rowStart = y * width;
            int run = 0;
            for (int x = 0; x < width; x++)
            {
                if (source[rowStart + x]) run = radius + 1;
                if (run > 0) { horizontal[rowStart + x] = true; run--; }
            }
            run = 0;
            for (int x = width - 1; x >= 0; x--)
            {
                if (source[rowStart + x]) run = radius + 1;
                if (run > 0) { horizontal[rowStart + x] = true; run--; }
            }
        }
        bool[] result = new bool[source.Length];
        for (int x = 0; x < width; x++)
        {
            int run = 0;
            for (int y = 0; y < height; y++)
            {
                int index = y * width + x;
                if (horizontal[index]) run = radius + 1;
                if (run > 0) { result[index] = true; run--; }
            }
            run = 0;
            for (int y = height - 1; y >= 0; y--)
            {
                int index = y * width + x;
                if (horizontal[index]) run = radius + 1;
                if (run > 0) { result[index] = true; run--; }
            }
        }
        return result;
    }

    private static int PixelDifference(byte[] before, byte[] after, int beforeIndex, int afterIndex)
    {
        int bp = beforeIndex * 3;
        int ap = afterIndex * 3;
        return Math.Max(
            Math.Abs(before[bp] - after[ap]),
            Math.Max(Math.Abs(before[bp + 1] - after[ap + 1]), Math.Abs(before[bp + 2] - after[ap + 2])));
    }

    private static void FindBestOffset(byte[] before, byte[] after, int width, int height, out int bestDx, out int bestDy)
    {
        bestDx = 0;
        bestDy = 0;
        long bestScore = long.MaxValue;
        int step = Math.Max(6, Math.Min(width, height) / 250);
        int border = MaximumAlignmentOffset + step;
        for (int dy = -MaximumAlignmentOffset; dy <= MaximumAlignmentOffset; dy++)
        {
            for (int dx = -MaximumAlignmentOffset; dx <= MaximumAlignmentOffset; dx++)
            {
                long score = 0;
                int samples = 0;
                for (int y = border; y < height - border; y += step)
                {
                    int by = y + dy;
                    for (int x = border; x < width - border; x += step)
                    {
                        int bx = x + dx;
                        int diff = PixelDifference(before, after, by * width + bx, y * width + x);
                        score += Math.Min(64, diff);
                        samples++;
                    }
                }
                if (samples > 0) score = (score * 1000L) / samples;
                score += (Math.Abs(dx) + Math.Abs(dy)) * 2L;
                if (score < bestScore)
                {
                    bestScore = score;
                    bestDx = dx;
                    bestDy = dy;
                }
            }
        }
    }

    private static byte[] AlignBefore(byte[] before, int width, int height, int dx, int dy)
    {
        if (dx == 0 && dy == 0) return before;
        byte[] aligned = new byte[before.Length];
        for (int i = 0; i < aligned.Length; i++) aligned[i] = 255;
        int copyPixels = width - Math.Abs(dx);
        if (copyPixels <= 0) return aligned;
        int copyBytes = copyPixels * 3;
        int sourceX = Math.Max(0, dx);
        int targetX = Math.Max(0, -dx);
        for (int y = 0; y < height; y++)
        {
            int sourceY = y + dy;
            if (sourceY < 0 || sourceY >= height) continue;
            Buffer.BlockCopy(before, (sourceY * width + sourceX) * 3, aligned, (y * width + targetX) * 3, copyBytes);
        }
        return aligned;
    }

    private static bool RegionsAreNear(PixelRegion a, PixelRegion b, int gap)
    {
        return a.minX <= b.maxX + gap && b.minX <= a.maxX + gap &&
               a.minY <= b.maxY + gap && b.minY <= a.maxY + gap;
    }

    private static void FinalizeRegion(PixelRegion region)
    {
        if (region.added >= region.count * 0.65) region.kind = "added";
        else if (region.removed >= region.count * 0.65) region.kind = "removed";
        else region.kind = "modified";
        double mean = region.count == 0 ? 0 : (double)region.diffTotal / region.count;
        region.confidence = Math.Max(0.45, Math.Min(0.99, 0.45 + (mean / 255.0) * 0.65));
    }

    private static List<PixelRegion> MergeNearbyRegions(List<PixelRegion> source, int gap)
    {
        int count = source.Count;
        if (count < 2) return source;

        int[] parent = new int[count];
        int[] rank = new int[count];
        for (int i = 0; i < count; i++) parent[i] = i;

        // Merge in one bounded pairwise pass. The previous remove-and-rescan loop
        // repeatedly restarted after every match and became cubic on noisy pages.
        for (int i = 0; i < count; i++)
        {
            for (int j = i + 1; j < count; j++)
            {
                if (!RegionsAreNear(source[i], source[j], gap)) continue;
                int rootA = FindRoot(parent, i);
                int rootB = FindRoot(parent, j);
                if (rootA == rootB) continue;
                if (rank[rootA] < rank[rootB]) parent[rootA] = rootB;
                else if (rank[rootA] > rank[rootB]) parent[rootB] = rootA;
                else { parent[rootB] = rootA; rank[rootA]++; }
            }
        }

        Dictionary<int, PixelRegion> byRoot = new Dictionary<int, PixelRegion>();
        for (int i = 0; i < count; i++)
        {
            int root = FindRoot(parent, i);
            PixelRegion target;
            if (!byRoot.TryGetValue(root, out target))
            {
                PixelRegion value = source[i];
                target = new PixelRegion
                {
                    minX = value.minX,
                    minY = value.minY,
                    maxX = value.maxX,
                    maxY = value.maxY,
                    count = value.count,
                    added = value.added,
                    removed = value.removed,
                    modified = value.modified,
                    diffTotal = value.diffTotal
                };
                byRoot[root] = target;
            }
            else
            {
                PixelRegion value = source[i];
                target.minX = Math.Min(target.minX, value.minX);
                target.minY = Math.Min(target.minY, value.minY);
                target.maxX = Math.Max(target.maxX, value.maxX);
                target.maxY = Math.Max(target.maxY, value.maxY);
                target.count += value.count;
                target.added += value.added;
                target.removed += value.removed;
                target.modified += value.modified;
                target.diffTotal += value.diffTotal;
            }
        }

        List<PixelRegion> merged = new List<PixelRegion>(byRoot.Values);
        for (int i = 0; i < merged.Count; i++) FinalizeRegion(merged[i]);
        return merged;
    }

    private static int FindRoot(int[] parent, int value)
    {
        int root = value;
        while (parent[root] != root) root = parent[root];
        while (parent[value] != value)
        {
            int next = parent[value];
            parent[value] = root;
            value = next;
        }
        return root;
    }

    private static List<PixelRegion> FindRegions(
        byte[] before,
        byte[] after,
        int width,
        int height,
        int threshold,
        int minimumRegionPixels,
        int padding,
        bool pageSizeChanged,
        out double changedRatio,
        out double averageDifference)
    {
        int pixels = width * height;
        int alignmentX;
        int alignmentY;
        FindBestOffset(before, after, width, height, out alignmentX, out alignmentY);
        before = AlignBefore(before, width, height, alignmentX, alignmentY);
        bool[] raw = new bool[pixels];
        byte[] pixelKind = new byte[pixels];
        int rawCount = 0;
        long rawDiff = 0;
        for (int i = 0; i < pixels; i++)
        {
            int p = i * 3;
            int db = Math.Abs(before[p] - after[p]);
            int dg = Math.Abs(before[p + 1] - after[p + 1]);
            int dr = Math.Abs(before[p + 2] - after[p + 2]);
            int max = Math.Max(db, Math.Max(dg, dr));
            if (max < threshold) continue;
            raw[i] = true;
            rawCount++;
            rawDiff += max;
            bool beforeWhite = IsWhite(before[p], before[p + 1], before[p + 2]);
            bool afterWhite = IsWhite(after[p], after[p + 1], after[p + 2]);
            pixelKind[i] = beforeWhite && !afterWhite ? (byte)1 : (!beforeWhite && afterWhite ? (byte)2 : (byte)3);
        }
        changedRatio = pixels == 0 ? 0 : (double)rawCount / pixels;
        averageDifference = rawCount == 0 ? 0 : (double)rawDiff / rawCount;
        if (rawCount == 0) return new List<PixelRegion>();
        if (!pageSizeChanged && changedRatio > 0.28 && averageDifference < 36)
            return new List<PixelRegion>();

        // Count changed neighbours with a summed-area table. The previous nested
        // 5x5 scan performed up to 25 lookups for every changed pixel and dominated
        // pages with widespread anti-aliasing differences.
        int integralStride = width + 1;
        int[] changedIntegral = new int[(height + 1) * integralStride];
        for (int y = 0; y < height; y++)
        {
            int rowTotal = 0;
            for (int x = 0; x < width; x++)
            {
                if (raw[y * width + x]) rowTotal++;
                changedIntegral[(y + 1) * integralStride + x + 1] =
                    changedIntegral[y * integralStride + x + 1] + rowTotal;
            }
        }
        bool[] candidate = new bool[pixels];
        for (int y = 0; y < height; y++)
        {
            int top = Math.Max(0, y - 2);
            int bottom = Math.Min(height - 1, y + 2) + 1;
            for (int x = 0; x < width; x++)
            {
                int index = y * width + x;
                if (!raw[index]) continue;
                int left = Math.Max(0, x - 2);
                int right = Math.Min(width - 1, x + 2) + 1;
                int neighbors =
                    changedIntegral[bottom * integralStride + right] -
                    changedIntegral[top * integralStride + right] -
                    changedIntegral[bottom * integralStride + left] +
                    changedIntegral[top * integralStride + left];
                candidate[index] = neighbors >= 3;
            }
        }

        // Group change pixels that are within 2x padding of each other. Without this the
        // connected-component pass returns one region per glyph, so a single edited number
        // produced five overlapping "M" boxes and a shifted row produced thousands.
        bool[] grouped = Dilate(candidate, width, height, padding);
        bool[] visited = new bool[pixels];
        int[] queue = new int[pixels];
        List<PixelRegion> regions = new List<PixelRegion>();
        for (int seed = 0; seed < pixels; seed++)
        {
            if (!grouped[seed] || visited[seed]) continue;
            int head = 0;
            int tail = 0;
            queue[tail++] = seed;
            visited[seed] = true;
            PixelRegion region = new PixelRegion
            {
                minX = int.MaxValue,
                maxX = -1,
                minY = int.MaxValue,
                maxY = -1
            };
            while (head < tail)
            {
                int index = queue[head++];
                int x = index % width;
                int y = index / width;
                if (candidate[index])
                {
                    region.count++;
                    if (x < region.minX) region.minX = x;
                    if (x > region.maxX) region.maxX = x;
                    if (y < region.minY) region.minY = y;
                    if (y > region.maxY) region.maxY = y;
                    if (pixelKind[index] == 1) region.added++;
                    else if (pixelKind[index] == 2) region.removed++;
                    else region.modified++;
                    int p = index * 3;
                    region.diffTotal += Math.Max(
                        Math.Abs(before[p] - after[p]),
                        Math.Max(Math.Abs(before[p + 1] - after[p + 1]), Math.Abs(before[p + 2] - after[p + 2])));
                }

                for (int yy = Math.Max(0, y - 1); yy <= Math.Min(height - 1, y + 1); yy++)
                {
                    for (int xx = Math.Max(0, x - 1); xx <= Math.Min(width - 1, x + 1); xx++)
                    {
                        int next = yy * width + xx;
                        if (!grouped[next] || visited[next]) continue;
                        visited[next] = true;
                        queue[tail++] = next;
                    }
                }
            }
            if (region.count < minimumRegionPixels || region.maxX < 0) continue;
            FinalizeRegion(region);
            region.minX = Math.Max(0, region.minX - padding);
            region.minY = Math.Max(0, region.minY - padding);
            region.maxX = Math.Min(width - 1, region.maxX + padding);
            region.maxY = Math.Min(height - 1, region.maxY + padding);
            regions.Add(region);
        }
        // The caller will switch to the safer side-by-side view when this many
        // disconnected candidates exist. Avoid quadratic merging work first.
        if (regions.Count > MaximumRegionsPerPage * 3) return regions;
        regions = MergeNearbyRegions(regions, Math.Max(padding * 2, Math.Min(width, height) / 130));
        regions.Sort(delegate (PixelRegion a, PixelRegion b)
        {
            int byY = a.minY.CompareTo(b.minY);
            return byY != 0 ? byY : a.minX.CompareTo(b.minX);
        });
        return regions;
    }

    private static void SaveBaseImage(string sourcePath, Bitmap source, Bitmap normalized, string destination)
    {
        if (!String.IsNullOrWhiteSpace(sourcePath) && File.Exists(sourcePath) &&
            source != null && source.Width == normalized.Width && source.Height == normalized.Height)
        {
            File.Copy(sourcePath, destination, true);
            return;
        }
        normalized.Save(destination, ImageFormat.Png);
    }

    private static bool TryGetImageSize(string path, out int width, out int height)
    {
        width = 0;
        height = 0;
        if (String.IsNullOrWhiteSpace(path) || !File.Exists(path)) return false;
        using (Image image = Image.FromFile(path))
        {
            width = image.Width;
            height = image.Height;
        }
        return width > 0 && height > 0;
    }

    private static string PrepareBaseAsset(
        string sourcePath,
        string outputDirectory,
        int pageNumber,
        bool before,
        bool copy,
        int width,
        int height)
    {
        if (!copy) return "render.png";
        string name = pageNumber.ToString("0000") + (before ? "-b.png" : "-a.png");
        string destination = Path.Combine(outputDirectory, name);
        if (!String.IsNullOrWhiteSpace(sourcePath) && File.Exists(sourcePath))
        {
            File.Copy(sourcePath, destination, true);
            return name;
        }
        using (Bitmap blank = new Bitmap(width, height, PixelFormat.Format24bppRgb))
        using (Graphics graphics = Graphics.FromImage(blank))
        {
            graphics.Clear(Color.White);
            blank.Save(destination, ImageFormat.Png);
        }
        return name;
    }

    private static ReportBinderDiffPage BuildSimplePage(
        string beforePath,
        string afterPath,
        string outputDirectory,
        int pageNumber,
        string forcedKind,
        bool copyBefore,
        bool copyAfter)
    {
        int beforeWidth;
        int beforeHeight;
        int afterWidth;
        int afterHeight;
        bool hasBefore = TryGetImageSize(beforePath, out beforeWidth, out beforeHeight);
        bool hasAfter = TryGetImageSize(afterPath, out afterWidth, out afterHeight);
        int width = Math.Max(hasBefore ? beforeWidth : 0, hasAfter ? afterWidth : 0);
        int height = Math.Max(hasBefore ? beforeHeight : 0, hasAfter ? afterHeight : 0);
        if (width <= 0 || height <= 0) return null;
        if (forcedKind == "unchanged" &&
            (!hasBefore || !hasAfter || beforeWidth != afterWidth || beforeHeight != afterHeight))
            return null;

        string beforeName = PrepareBaseAsset(
            beforePath, outputDirectory, pageNumber, true, copyBefore, width, height);
        string afterName = PrepareBaseAsset(
            afterPath, outputDirectory, pageNumber, false, copyAfter, width, height);
        List<ReportBinderDiffRegion> regions = new List<ReportBinderDiffRegion>();
        if (forcedKind == "added" || forcedKind == "removed")
        {
            int inset = Math.Max(6, Math.Min(width, height) / 250);
            regions.Add(new ReportBinderDiffRegion
            {
                regionId = "p" + pageNumber.ToString("0000") + "-r0001",
                kind = forcedKind,
                x = (double)inset / width,
                y = (double)inset / height,
                width = (double)Math.Max(1, width - (inset * 2)) / width,
                height = (double)Math.Max(1, height - (inset * 2)) / height,
                confidence = 1,
                pixelCount = width * height
            });
        }
        bool unknown = forcedKind == "unknown";
        return new ReportBinderDiffPage
        {
            pageNumber = pageNumber,
            width = width,
            height = height,
            pageSizeChanged = hasBefore && hasAfter &&
                (beforeWidth != afterWidth || beforeHeight != afterHeight),
            status = unknown ? "unknown" : "ready",
            message = unknown ? "信頼できる差分領域を判定できません。" : "",
            confidence = unknown ? 0 : 1,
            changedRatio = (forcedKind == "added" || forcedKind == "removed") ? 1 : 0,
            regionCount = regions.Count,
            regions = regions.ToArray(),
            beforeFile = beforeName,
            afterFile = afterName,
            beforeMaskFile = "",
            beforeOverlayFile = "",
            afterMaskFile = "",
            afterOverlayFile = ""
        };
    }

    public static ReportBinderDiffPage ComparePage(
        string beforePath,
        string afterPath,
        string outputDirectory,
        int pageNumber,
        string forcedKind,
        int threshold,
        int minimumRegionPixels,
        int padding)
    {
        return ComparePage(beforePath, afterPath, outputDirectory, pageNumber, forcedKind,
            threshold, minimumRegionPixels, padding, true, true);
    }

    private static ReportBinderDiffPage ComparePage(
        string beforePath,
        string afterPath,
        string outputDirectory,
        int pageNumber,
        string forcedKind,
        int threshold,
        int minimumRegionPixels,
        int padding,
        bool copyBefore,
        bool copyAfter)
    {
        Directory.CreateDirectory(outputDirectory);
        if (forcedKind == "unchanged" || forcedKind == "added" ||
            forcedKind == "removed" || forcedKind == "unknown")
        {
            ReportBinderDiffPage simple = BuildSimplePage(
                beforePath, afterPath, outputDirectory, pageNumber, forcedKind, copyBefore, copyAfter);
            if (simple != null) return simple;
        }
        Bitmap beforeSource = LoadImage(beforePath);
        Bitmap afterSource = LoadImage(afterPath);
        try
        {
            int width = Math.Max(beforeSource == null ? 0 : beforeSource.Width, afterSource == null ? 0 : afterSource.Width);
            int height = Math.Max(beforeSource == null ? 0 : beforeSource.Height, afterSource == null ? 0 : afterSource.Height);
            if (width <= 0) width = 1240;
            if (height <= 0) height = 1754;
            bool sizeChanged = beforeSource != null && afterSource != null &&
                (beforeSource.Width != afterSource.Width || beforeSource.Height != afterSource.Height);
            using (Bitmap before = Normalize(beforeSource, width, height))
            using (Bitmap after = Normalize(afterSource, width, height))
            {
                string stem = pageNumber.ToString("0000");
                string beforeName = copyBefore ? stem + "-b.png" : "render.png";
                string afterName = copyAfter ? stem + "-a.png" : "render.png";
                if (copyBefore)
                    SaveBaseImage(beforePath, beforeSource, before, Path.Combine(outputDirectory, beforeName));
                if (copyAfter)
                    SaveBaseImage(afterPath, afterSource, after, Path.Combine(outputDirectory, afterName));

                List<PixelRegion> pixelRegions = new List<PixelRegion>();
                double changedRatio = 0;
                double averageDifference = 0;
                string status = "ready";
                string message = "";

                if (forcedKind == "added")
                {
                    pixelRegions.Add(FullPageRegion(width, height, "added"));
                    changedRatio = 1;
                }
                else if (forcedKind == "removed")
                {
                    pixelRegions.Add(FullPageRegion(width, height, "removed"));
                    changedRatio = 1;
                }
                else if (forcedKind == "unknown")
                {
                    status = "unknown";
                    message = "信頼できる差分領域を判定できません。";
                }
                else
                {
                    pixelRegions = FindRegions(
                        ReadBgr(before),
                        ReadBgr(after),
                        width,
                        height,
                        Math.Max(1, threshold),
                        Math.Max(1, minimumRegionPixels),
                        Math.Max(0, padding),
                        sizeChanged,
                        out changedRatio,
                        out averageDifference);
                    bool sparseFullPage = false;
                    foreach (PixelRegion region in pixelRegions)
                    {
                        double area = (double)(region.maxX - region.minX + 1) * (region.maxY - region.minY + 1);
                        double areaRatio = area / Math.Max(1.0, (double)width * height);
                        double density = region.count / Math.Max(1.0, area);
                        if (areaRatio > 0.72 && density < 0.08) { sparseFullPage = true; break; }
                    }
                    if ((!sizeChanged && changedRatio > 0.28 && averageDifference < 36) || sparseFullPage)
                    {
                        pixelRegions.Clear();
                        status = "unknown";
                        message = "微小な描画差がページ全体へ広がっているため、誤解を招く全ページ強調を停止しました。";
                    }
                    else if (pixelRegions.Count > MaximumRegionsPerPage)
                    {
                        // A shifted row can change almost every glyph on the page. Drawing that many
                        // boxes makes the overlay unreadable and the region navigation unusable.
                        pixelRegions.Clear();
                        status = "unknown";
                        message = "変更領域が多すぎるため、領域の強調を停止しました。左右の表示で確認してください。";
                    }
                }

                // Region rectangles and labels are rendered by the browser from the JSON below.
                // Avoid four full-page PNG encodes (before/after mask + overlay) per page.
                string beforeMask = "";
                string beforeOverlay = "";
                string afterMask = "";
                string afterOverlay = "";

                List<ReportBinderDiffRegion> publicRegions = new List<ReportBinderDiffRegion>();
                double confidenceTotal = 0;
                for (int i = 0; i < pixelRegions.Count; i++)
                {
                    PixelRegion region = pixelRegions[i];
                    confidenceTotal += region.confidence;
                    publicRegions.Add(new ReportBinderDiffRegion
                    {
                        regionId = "p" + pageNumber.ToString("0000") + "-r" + (i + 1).ToString("0000"),
                        kind = region.kind,
                        x = (double)region.minX / width,
                        y = (double)region.minY / height,
                        width = (double)(region.maxX - region.minX + 1) / width,
                        height = (double)(region.maxY - region.minY + 1) / height,
                        confidence = region.confidence,
                        pixelCount = region.count
                    });
                }

                return new ReportBinderDiffPage
                {
                    pageNumber = pageNumber,
                    width = width,
                    height = height,
                    pageSizeChanged = sizeChanged,
                    status = status,
                    message = message,
                    confidence = publicRegions.Count == 0 ? (status == "unknown" ? 0 : 1) : confidenceTotal / publicRegions.Count,
                    changedRatio = changedRatio,
                    regionCount = publicRegions.Count,
                    regions = publicRegions.ToArray(),
                    beforeFile = beforeName,
                    afterFile = afterName,
                    beforeMaskFile = beforeMask,
                    beforeOverlayFile = beforeOverlay,
                    afterMaskFile = afterMask,
                    afterOverlayFile = afterOverlay
                };
            }
        }
        finally
        {
            if (beforeSource != null) beforeSource.Dispose();
            if (afterSource != null) afterSource.Dispose();
        }
    }

    public static ReportBinderDiffBatchPageResult[] ComparePages(
        ReportBinderDiffBatchPageRequest[] requests,
        int maximumDegreeOfParallelism,
        int threshold,
        int minimumRegionPixels,
        int padding)
    {
        if (requests == null || requests.Length == 0)
            return new ReportBinderDiffBatchPageResult[0];

        ReportBinderDiffBatchPageResult[] results = new ReportBinderDiffBatchPageResult[requests.Length];
        ParallelOptions options = new ParallelOptions
        {
            MaxDegreeOfParallelism = Math.Max(1, Math.Min(4, maximumDegreeOfParallelism))
        };
        Parallel.For(0, requests.Length, options, delegate (int index)
        {
            ReportBinderDiffBatchPageRequest request = requests[index];
            ReportBinderDiffBatchPageResult result = new ReportBinderDiffBatchPageResult
            {
                itemId = request == null ? "" : request.itemId,
                pageNumber = request == null ? 0 : request.pageNumber,
                error = ""
            };
            try
            {
                if (request == null) throw new ArgumentNullException("request");
                result.page = ComparePage(
                    request.beforePath,
                    request.afterPath,
                    request.outputDirectory,
                    request.pageNumber,
                    request.kind,
                    threshold,
                    minimumRegionPixels,
                    padding,
                    request.copyBefore,
                    request.copyAfter);
            }
            catch (Exception ex)
            {
                result.error = ex.Message;
            }
            results[index] = result;
        });
        return results;
    }
}
