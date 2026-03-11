using iText.Kernel.Colors;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using Spire.Pdf.Conversion;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
                OfdConverter converter = new OfdConverter(Input);
                converter.ToPdf(OutPut);
                RemoveEvaluationWarning(OutPut);
                return ConvertResult.Successful;
            }
            catch (Exception)
            {
                return ConvertResult.Failed;
            }
        }

        private static void RemoveEvaluationWarning(string pdfPath)
        {
            var tempPath = pdfPath + ".clean";
            try
            {
                using (var reader = new PdfReader(pdfPath))
                using (var writer = new PdfWriter(tempPath))
                using (var doc = new PdfDocument(reader, writer))
                {
                    // Spire.PDF (FreeSpire) renders a ~30-point-tall evaluation warning strip
                    // at the top of every page as part of the page content stream.
                    // Cover it by drawing a white rectangle on top of each page.
                    const float warningHeight = 30f;
                    for (int i = 1; i <= doc.GetNumberOfPages(); i++)
                    {
                        var page = doc.GetPage(i);
                        var pageSize = page.GetPageSize();
                        var canvas = new PdfCanvas(page.NewContentStreamAfter(), page.GetResources(), doc);
                        canvas.SaveState()
                              .SetFillColor(ColorConstants.WHITE)
                              .Rectangle(pageSize.GetLeft(), pageSize.GetTop() - warningHeight,
                                         pageSize.GetWidth(), warningHeight)
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
    }
}
