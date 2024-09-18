using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using KeyStone.PoC.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

namespace KeyStone.PoC.Controllers;
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

    public IActionResult Privacy()
    {
        return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }

    public IActionResult RichEdit()
    {
        return View();
    }
    /*
    [HttpPost("CalculateDocumentVariable")]
    public IActionResult CalculateDocumentVariable([FromForm] string variableName, [FromForm] string arguments)
    {
        RichEditDocumentServer documentServer = new RichEditDocumentServer();
        // Access the document
        Document document = documentServer.Document;

        // Specify the position where the table will be inserted
        DocumentPosition position = document.Range.End;

        // Create a table with 3 rows and 4 columns at the specified position
        Table table = document.Tables.Create(position, 3, 4);

        // Iterate through the rows and cells to populate the table
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            TableRow row = table.Rows[rowIndex];
            for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
            {
                TableCell cell = row.Cells[colIndex];
                // Begin updating the cell
                SubDocument cellSubDoc = cell.Range.BeginUpdateDocument();
                // Insert text into the cell
                cellSubDoc.InsertText(cellSubDoc.CreatePosition(cellSubDoc.Range.End.ToInt()), $"{variableName} {rowIndex + 1},{colIndex + 1}");
                // End updating the cell
                cell.Range.EndUpdateDocument(cellSubDoc);
            }
        }
        // Get the RTF content of the document
        string rtfContent = document.GetRtfText(document.Range);

        // Return the RTF content
        return Content(rtfContent, "text/plain");
    }*/
}

