using iText.Kernel.Colors;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using Microsoft.AspNetCore.Mvc;
using Spire.Pdf.Conversion;

namespace Ofd2PdfService.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ConvertController : ControllerBase
{
    private readonly ILogger<ConvertController> _logger;

    public ConvertController(ILogger<ConvertController> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Convert an OFD file to PDF.
    /// </summary>
    /// <remarks>
    /// POST /api/convert
    /// Content-Type: multipart/form-data
    /// Body: file=&lt;OFD file&gt;
    ///
    /// Returns the converted PDF as a file download.
    /// </remarks>
    [HttpPost]
    [Consumes("multipart/form-data")]
    [ProducesResponseType(typeof(FileContentResult), StatusCodes.Status200OK)]
    [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status400BadRequest)]
    [ProducesResponseType(typeof(ProblemDetails), StatusCodes.Status500InternalServerError)]
    public async Task<IActionResult> Post([FromForm] IFormFile file)
    {
        if (file == null || file.Length == 0)
        {
            return BadRequest(new ProblemDetails
            {
                Title = "No file provided",
                Detail = "Please upload an OFD file using the 'file' form field.",
                Status = StatusCodes.Status400BadRequest
            });
        }

        var extension = Path.GetExtension(file.FileName);
        if (!string.Equals(extension, ".ofd", StringComparison.OrdinalIgnoreCase))
        {
            return BadRequest(new ProblemDetails
            {
                Title = "Invalid file type",
                Detail = $"Expected an .ofd file but received '{extension}'.",
                Status = StatusCodes.Status400BadRequest
            });
        }

        var tmpDir = Path.Combine(Path.GetTempPath(), "ofd2pdf", Guid.NewGuid().ToString());
        Directory.CreateDirectory(tmpDir);

        // Use a safe, generated filename to avoid path traversal from user-supplied names
        var safeInputName = Path.GetRandomFileName() + ".ofd";
        var inputPath = Path.Combine(tmpDir, safeInputName);
        var outputPath = Path.ChangeExtension(inputPath, ".pdf");

        try
        {
            using (var stream = new FileStream(inputPath, FileMode.Create, FileAccess.Write))
            {
                await file.CopyToAsync(stream);
            }

            _logger.LogInformation("Converting {FileName} to PDF", file.FileName);
            var converter = new OfdConverter(inputPath);
            converter.ToPdf(outputPath);
            RemoveEvaluationWarning(outputPath, _logger);

            var pdfBytes = await System.IO.File.ReadAllBytesAsync(outputPath);
            var pdfFileName = Path.ChangeExtension(Path.GetFileName(file.FileName), ".pdf");

            _logger.LogInformation("Conversion succeeded: {FileName}", pdfFileName);
            return File(pdfBytes, "application/pdf", pdfFileName);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Conversion failed for {FileName}", file.FileName);
            return StatusCode(StatusCodes.Status500InternalServerError, new ProblemDetails
            {
                Title = "Conversion failed",
                Detail = ex.Message,
                Status = StatusCodes.Status500InternalServerError
            });
        }
        finally
        {
            try { Directory.Delete(tmpDir, recursive: true); }
            catch (Exception ex) { _logger.LogWarning(ex, "Failed to clean up temp directory {TmpDir}", tmpDir); }
        }
    }

    /// <summary>
    /// Removes the Spire.PDF evaluation warning from all pages of a PDF file by painting a white
    /// rectangle over the top of each page, covering the warning strip in the content stream.
    /// </summary>
    private static void RemoveEvaluationWarning(string pdfPath, ILogger logger)
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
            System.IO.File.Move(tempPath, pdfPath, overwrite: true);
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "Failed to remove Spire.PDF evaluation warning from {PdfPath}; the original PDF will be used as-is.", pdfPath);
            if (System.IO.File.Exists(tempPath))
                System.IO.File.Delete(tempPath);
        }
    }
}
