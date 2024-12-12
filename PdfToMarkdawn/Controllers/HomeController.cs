using Microsoft.AspNetCore.Mvc;
using Microsoft.VisualBasic;
using PdfToMarkdawn.Models;
using System.Diagnostics;
using System.Reflection.Metadata;
using Aspose.Words;
using Document = Aspose.Words.Document;
using Westwind.AspNetCore.Markdown;
using Microsoft.AspNetCore.Http;
namespace PdfToMarkdawn.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private IWebHostEnvironment _webHostEnvironment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment webHostEnvironment)
        {
            _logger = logger;
            _webHostEnvironment = webHostEnvironment;
        }

        public IActionResult Index()
        {
            TestModel test = new TestModel();


            var fileDic = "Files";
            string filePath = Path.Combine(_webHostEnvironment.WebRootPath, fileDic);

            if (!Directory.Exists(filePath))
                Directory.CreateDirectory(filePath);

            filePath = Path.Combine(filePath, "output.md");

            test.MarkdownText = Markdown.ParseHtmlStringFromFile(filePath);
            return View(test);
        }

        public IActionResult Upload(IFormFile file)
        {
            var stream = file.OpenReadStream();

            var fileDic = "Files";
            string filePath = Path.Combine(_webHostEnvironment.WebRootPath, fileDic);

            if (!Directory.Exists(filePath))
                Directory.CreateDirectory(filePath);

            var fileName = file.FileName;
            filePath = Path.Combine(filePath, "output.md");
            Document document = new Document(stream);
            document.Save("DocOutput.doc", SaveFormat.Doc);
            var outputDocument = new Aspose.Words.Document("DocOutput.doc");
            outputDocument.Save(filePath, SaveFormat.Markdown);
            TestModel test = new TestModel();
            test.MarkdownText = Markdown.ParseHtmlStringFromFile(filePath);

            return View("Index", test);
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
