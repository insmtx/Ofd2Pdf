using iText.Kernel.Colors;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using Spire.Pdf.Conversion;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace Ofd2Pdf
{
    public enum ConvertResult
    {
        Successful,
        Failed
    }
    public class Converter
    {
        public ConvertResult ConvertToPdf(string Input, string OutPut)
        {
            Console.WriteLine(Input + " " + OutPut);
            if (Input == null || OutPut == null)
            {
                return ConvertResult.Failed;
            }

            if (!File.Exists(Input))
            {
                return ConvertResult.Failed;
            }

            try
            {
                ConvertOfdToPdf(Input, OutPut);
                RemoveEvaluationWarning(OutPut);
                return ConvertResult.Successful;
            }
            catch (Exception)
            {
                return ConvertResult.Failed;
            }
        }

        /// <summary>
        /// Converts an OFD file to PDF, splitting it into batches of up to 3 pages to work
        /// around the FreeSpire.PDF evaluation version page limit, then merging the results.
        /// </summary>
        private static void ConvertOfdToPdf(string inputPath, string outputPath)
        {
            int totalPages;
            try { totalPages = GetOfdPageCount(inputPath); }
            catch { totalPages = 0; }

            if (totalPages <= 3)
            {
                new OfdConverter(inputPath).ToPdf(outputPath);
                return;
            }

            string tmpDir = Path.Combine(Path.GetTempPath(), "ofd2pdf_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tmpDir);
            try
            {
                List<string> partPdfs = new List<string>();
                for (int start = 0; start < totalPages; start += 3)
                {
                    string partOfd = Path.Combine(tmpDir, "part_" + start + ".ofd");
                    string partPdf = Path.Combine(tmpDir, "part_" + start + ".pdf");
                    CreateSubOfd(inputPath, partOfd, start, Math.Min(3, totalPages - start));
                    new OfdConverter(partOfd).ToPdf(partPdf);
                    partPdfs.Add(partPdf);
                }
                MergePdfs(partPdfs, outputPath);
            }
            finally
            {
                try { Directory.Delete(tmpDir, recursive: true); } catch { }
            }
        }

        /// <summary>
        /// Returns the total number of pages in an OFD file by reading its Document.xml.
        /// </summary>
        private static int GetOfdPageCount(string ofdPath)
        {
            using (ZipArchive archive = ZipFile.OpenRead(ofdPath))
            {
                ZipArchiveEntry ofdXmlEntry = null;
                foreach (ZipArchiveEntry e in archive.Entries)
                    if (e.FullName.Equals("OFD.xml", StringComparison.OrdinalIgnoreCase)) { ofdXmlEntry = e; break; }
                if (ofdXmlEntry == null) return 0;

                string docRootPath;
                using (Stream stream = ofdXmlEntry.Open())
                    docRootPath = XDocument.Load(stream).Descendants()
                        .First(e => e.Name.LocalName == "DocRoot")
                        .Value.Trim().Replace('\\', '/');

                ZipArchiveEntry docEntry = null;
                foreach (ZipArchiveEntry e in archive.Entries)
                    if (e.FullName.Replace('\\', '/').Equals(docRootPath, StringComparison.OrdinalIgnoreCase)) { docEntry = e; break; }
                if (docEntry == null) return 0;

                using (Stream docStream = docEntry.Open())
                    return XDocument.Load(docStream).Descendants()
                        .Count(e => e.Name.LocalName == "Page" && e.Attribute("BaseLoc") != null);
            }
        }

        /// <summary>
        /// Creates a sub-OFD ZIP file containing only the specified page range.
        /// All non-page entries (resources, fonts, etc.) are copied as-is so that
        /// each sub-OFD is a valid, self-contained document.
        /// </summary>
        private static void CreateSubOfd(string sourcePath, string destPath, int startPage, int pageCount)
        {
            using (ZipArchive src = ZipFile.OpenRead(sourcePath))
            {
                ZipArchiveEntry ofdXmlEntry = null;
                foreach (ZipArchiveEntry e in src.Entries)
                    if (e.FullName.Equals("OFD.xml", StringComparison.OrdinalIgnoreCase)) { ofdXmlEntry = e; break; }

                string docRootPath;
                using (Stream stream = ofdXmlEntry.Open())
                    docRootPath = XDocument.Load(stream).Descendants()
                        .First(e => e.Name.LocalName == "DocRoot")
                        .Value.Trim().Replace('\\', '/');

                int lastSlash = docRootPath.LastIndexOf('/');
                string docFolder = lastSlash >= 0 ? docRootPath.Substring(0, lastSlash) : "";

                XDocument docXml;
                List<XElement> allPages;
                ZipArchiveEntry docEntry = null;
                foreach (ZipArchiveEntry e in src.Entries)
                    if (e.FullName.Replace('\\', '/').Equals(docRootPath, StringComparison.OrdinalIgnoreCase)) { docEntry = e; break; }
                using (Stream stream = docEntry.Open())
                {
                    docXml = XDocument.Load(stream);
                    allPages = docXml.Descendants()
                        .Where(e => e.Name.LocalName == "Page" && e.Attribute("BaseLoc") != null)
                        .ToList();
                }

                List<XElement> selectedPages = allPages.Skip(startPage).Take(pageCount).ToList();

                string pagesFolderPrefix = string.IsNullOrEmpty(docFolder) ? "Pages/" : docFolder + "/Pages/";
                HashSet<string> includedPrefixes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (XElement p in selectedPages)
                {
                    string baseLoc = p.Attribute("BaseLoc").Value.Replace('\\', '/').TrimStart('/');
                    includedPrefixes.Add((string.IsNullOrEmpty(docFolder) ? baseLoc : docFolder + "/" + baseLoc) + "/");
                }

                using (ZipArchive dest = ZipFile.Open(destPath, ZipArchiveMode.Create))
                {
                    foreach (ZipArchiveEntry entry in src.Entries)
                    {
                        string name = entry.FullName.Replace('\\', '/');
                        if (name.EndsWith("/")) continue;

                        if (name.StartsWith(pagesFolderPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            bool included = false;
                            foreach (string prefix in includedPrefixes)
                                if (name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) { included = true; break; }
                            if (!included) continue;
                        }

                        if (name.Equals(docRootPath, StringComparison.OrdinalIgnoreCase))
                        {
                            XDocument modifiedDoc = new XDocument(docXml);
                            XElement pagesEl = modifiedDoc.Descendants().First(e => e.Name.LocalName == "Pages");
                            pagesEl.RemoveNodes();
                            foreach (XElement page in selectedPages)
                                pagesEl.Add(new XElement(page));

                            using (Stream destStream = dest.CreateEntry(name).Open())
                                modifiedDoc.Save(destStream);
                        }
                        else
                        {
                            using (Stream destStream = dest.CreateEntry(name).Open())
                            using (Stream srcStream = entry.Open())
                                srcStream.CopyTo(destStream);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Merges multiple PDF files into a single output PDF using iText.
        /// </summary>
        private static void MergePdfs(List<string> pdfPaths, string outputPath)
        {
            List<PdfReader> readers = pdfPaths.Select(p => new PdfReader(p)).ToList();
            List<PdfDocument> srcDocs = readers.Select(r => new PdfDocument(r)).ToList();
            try
            {
                using (PdfWriter writer = new PdfWriter(outputPath))
                using (PdfDocument mergedDoc = new PdfDocument(writer))
                {
                    foreach (PdfDocument srcDoc in srcDocs)
                        srcDoc.CopyPagesTo(1, srcDoc.GetNumberOfPages(), mergedDoc);
                }
            }
            finally
            {
                foreach (PdfDocument doc in srcDocs)
                    try { doc.Close(); } catch { }
            }
        }

        private static void RemoveEvaluationWarning(string pdfPath)        {
            var tempPath = pdfPath + ".clean";
            try
            {
                using (var reader = new PdfReader(pdfPath))
                using (var writer = new PdfWriter(tempPath))
                using (var doc = new PdfDocument(reader, writer))
                {
                    for (int i = 1; i <= doc.GetNumberOfPages(); i++)
                    {
                        var page = doc.GetPage(i);
                        var warningRect = FindEvaluationWarningRect(page);
                        if (warningRect == null)
                            continue;

                        var canvas = new PdfCanvas(page.NewContentStreamAfter(), page.GetResources(), doc);
                        canvas.SaveState()
                              .SetFillColor(ColorConstants.WHITE)
                              .Rectangle(warningRect)
                              .Fill()
                              .RestoreState()
                              .Release();
                    }
                }
                // File.Replace atomically replaces the destination with the source on the same volume.
                System.IO.File.Replace(tempPath, pdfPath, destinationBackupFileName: null);
            }
            catch (Exception)
            {
                // Removal of the evaluation warning is best-effort; if it fails the original
                // converted PDF (with the watermark) is kept so the conversion still succeeds.
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        /// <summary>
        /// Parses the page content stream for text containing "Evaluation Warning" and returns
        /// a page-width rectangle that covers those text lines. Returns null if no match is found.
        /// </summary>
        private static Rectangle FindEvaluationWarningRect(PdfPage page)
        {
            var locator = new EvaluationWarningLocator();
            new PdfCanvasProcessor(locator).ProcessPageContent(page);
            return locator.GetCoveringRect(page.GetPageSize());
        }

        /// <summary>
        /// Collects TextRenderInfo events from the page content stream, groups them into
        /// text lines by baseline Y, and identifies lines containing "Evaluation Warning"
        /// for targeted removal.
        /// </summary>
        private sealed class EvaluationWarningLocator : IEventListener
        {
            private readonly List<TextChunkInfo> _chunks = new List<TextChunkInfo>();

            private struct TextChunkInfo
            {
                public string Text;
                public float BaselineY;
                public float Top;
                public float Bottom;
            }

            public void EventOccurred(IEventData data, EventType type)
            {
                var info = data as TextRenderInfo;
                if (info == null) return;
                var text = info.GetText();
                if (string.IsNullOrEmpty(text)) return;

                _chunks.Add(new TextChunkInfo
                {
                    Text = text,
                    BaselineY = info.GetBaseline().GetBoundingRectangle().GetBottom(),
                    Top = info.GetAscentLine().GetBoundingRectangle().GetTop(),
                    Bottom = info.GetDescentLine().GetBoundingRectangle().GetBottom(),
                });
            }

            public ICollection<EventType> GetSupportedEvents()
            {
                return new HashSet<EventType> { EventType.RENDER_TEXT };
            }

            /// <summary>
            /// Groups chunks by baseline Y, concatenates each line's text, and returns a
            /// page-width rectangle covering every line that contains "Evaluation Warning".
            /// Returns null if no match is found.
            /// </summary>
            public Rectangle GetCoveringRect(Rectangle pageSize)
            {
                const float yTolerance = 2f;
                const float margin = 2f;

                // Group chunks into lines by baseline Y bucket (O(n) via dictionary).
                var lines = new Dictionary<int, List<TextChunkInfo>>();
                foreach (var chunk in _chunks)
                {
                    int key = (int)Math.Round(chunk.BaselineY / yTolerance);
                    List<TextChunkInfo> line;
                    if (!lines.TryGetValue(key, out line))
                    {
                        line = new List<TextChunkInfo>();
                        lines[key] = line;
                    }
                    line.Add(chunk);
                }

                // Find lines whose concatenated text contains "Evaluation Warning".
                float minBottom = float.MaxValue;
                float maxTop = float.MinValue;
                bool found = false;

                foreach (var line in lines.Values)
                {
                    var lineText = string.Concat(line.Select(c => c.Text));
                    if (lineText.IndexOf("Evaluation Warning", StringComparison.OrdinalIgnoreCase) < 0)
                        continue;

                    found = true;
                    foreach (var chunk in line)
                    {
                        if (chunk.Bottom < minBottom) minBottom = chunk.Bottom;
                        if (chunk.Top > maxTop) maxTop = chunk.Top;
                    }
                }

                if (!found)
                    return null;

                float bottom = minBottom - margin;
                return new Rectangle(pageSize.GetLeft(), bottom, pageSize.GetWidth(), maxTop - bottom + margin);
            }
        }
    }
}
