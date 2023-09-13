using Microsoft.AspNetCore.Mvc;
using PPTtoPDF.Models;
using System.Diagnostics;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using Syncfusion.Pdf;

namespace PPTtoPDF.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult ConvertToPDF()
        {
            using(FileStream input = new FileStream(Path.GetFullPath("Data/Template.pptx"), FileMode.Open, FileAccess.Read))
            {
                using (IPresentation pptxDoc = Presentation.Open(input))
                {
                    PresentationToPdfConverterSettings settings = new PresentationToPdfConverterSettings();
                    //settings.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;
                    //settings.AutoTag = true;
                    settings.PublishOptions = PublishOptions.NotesPages;
                    PdfDocument pdfDocument = PresentationToPdfConverter.Convert(pptxDoc, settings);
                    MemoryStream ms = new MemoryStream();
                    pdfDocument.Save(ms);
                    ms.Position = 0;
                    return File(ms, "application/pdf", "PPTtoPDF.pdf");
                }
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}