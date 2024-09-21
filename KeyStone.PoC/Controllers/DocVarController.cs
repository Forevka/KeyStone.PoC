using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using Microsoft.AspNetCore.Mvc;
using System.Drawing;

namespace KeyStone.PoC.Controllers;

[ApiController]
[Route("api/[controller]")]
public class DocVarController : ControllerBase
{
    private readonly IWebHostEnvironment _webHostEnvironment;

    public DocVarController(IWebHostEnvironment webHostEnvironment)
    {
        _webHostEnvironment = webHostEnvironment;
    }

    [HttpPost("CalculateDocumentVariable")]
    public async Task<IActionResult> CalculateDocumentVariable([FromForm] string variableName, [FromForm] List<string>? arguments)
    {
        var document = variableName switch
        {
            "ComplexTable" => GenerateComplexTable(),
            "TableWithImages" => await GenerateTableWithImages(),
            "FormattedText" => GenerateFormattedText(),
            "CTable" => GenerateCTable(arguments ?? []),
            _ => GenerateComplexTable(),
        };

        using var docxStream = new MemoryStream();
        document.SaveDocument(docxStream, DocumentFormat.OpenXml);

        var base64Content = Convert.ToBase64String(docxStream.ToArray());
        return Content(base64Content, "text/plain");
    }

    private RichEditDocumentServer GenerateCTable(List<string> arguments)
    {
        var documentServer = new RichEditDocumentServer();

        var document = documentServer.Document;

        var scenarioValue = arguments.FirstOrDefault();

        var scenario = 0;
        if (!string.IsNullOrEmpty(scenarioValue))
            int.TryParse(scenarioValue, out scenario);

        List<Action<Document>> scenarios =
        [
            TableActions.CreateTable, //0
            TableActions.CreateFixedTable,
            TableActions.ChangeTableColor, //2
            TableActions.CreateAndApplyTableStyle,
            TableActions.ChangeColumnAppearance, //4
            TableActions.UseTableCellProcessor,
            TableActions.MergeCells, //6
            TableActions.SplitCells,
            GenerateCustomerTableWithLogo //8
        ];

        scenario = Math.Min(Math.Max(scenario, 0), scenarios.Count - 1);

        var scenarioAction = scenarios.ElementAtOrDefault(scenario) ?? scenarios[0];

        scenarioAction(document);

        return documentServer;
    }

    private void GenerateCustomerTableWithLogo(Document document)
    {
        var table = document.Tables.Create(document.Range.Start, 2, 4);

        document.InsertSingleLineText(table.Rows[0].Cells[0].Range.Start, "ID");
        document.InsertSingleLineText(table.Rows[0].Cells[1].Range.Start, "Photo");
        document.InsertSingleLineText(table.Rows[0].Cells[2].Range.Start, "Customer Info");
        document.InsertSingleLineText(table.Rows[0].Cells[3].Range.Start, "Rentals");

        for (var i = 1; i < 2; i++)
        {
            document.InsertSingleLineText(table.Rows[i].Cells[0].Range.Start, $"ID {i}");

            var customerInfo = $"Customer Info {i}\n" +
                               $"Address: 123 Main St, Apt {i}\n" +
                               $"Phone: (555) 123-456{i}\n" +
                               $"Email: customer{i}@example.com";
            document.InsertText(table.Rows[i].Cells[2].Range.Start, customerInfo);

            var rentalsInfo = $"Rental {i}\n" +
                              $"Date: 01/01/202{i}\n" +
                              $"Amount: ${100 * i}\n" +
                              $"Status: Active";
            document.InsertText(table.Rows[i].Cells[3].Range.Start, rentalsInfo);
        }

        for (var i = 1; i < 2; i++)
        {
            var imagePath = Path.Combine(_webHostEnvironment.ContentRootPath, "Documents", "logo.png");
            if (System.IO.File.Exists(imagePath))
            {
                using var imageStream = new FileStream(imagePath, FileMode.Open);

                var documentImageSource = DocumentImageSource.FromStream(imageStream);
                document.Images.Insert(table.Rows[i].Cells[1].Range.Start, documentImageSource);
            }
        }

        //document.SaveDocument(Path.Combine(_webHostEnvironment.ContentRootPath, "Documents", "GenerateCustomerTableWithLogo.docx"), DocumentFormat.OpenXml);
    }

    private RichEditDocumentServer GenerateFormattedText()
    {
        var documentServer = new RichEditDocumentServer();

        var document = documentServer.Document;

        // Insert a heading
        var headingRange = document.InsertText(document.Range.End, "This is a Heading\n");
        var headingProps = document.BeginUpdateCharacters(headingRange);
        headingProps.FontSize = 24;
        headingProps.Bold = true;
        headingProps.ForeColor = Color.DarkBlue;
        document.EndUpdateCharacters(headingProps);

        // Insert a paragraph with different formatting
        var paraRange = document.InsertText(document.Range.End, "This is a paragraph with some ");
        var paraProps = document.BeginUpdateCharacters(paraRange);
        paraProps.FontSize = 12;
        document.EndUpdateCharacters(paraProps);

        // Insert italic text
        var italicRange = document.InsertText(document.Range.End, "italic ");
        var italicProps = document.BeginUpdateCharacters(italicRange);
        italicProps.Italic = true;
        document.EndUpdateCharacters(italicProps);

        // Insert bold text
        var boldRange = document.InsertText(document.Range.End, "and bold ");
        var boldProps = document.BeginUpdateCharacters(boldRange);
        boldProps.Bold = true;
        document.EndUpdateCharacters(boldProps);

        // Insert underlined text
        var underlineRange = document.InsertText(document.Range.End, "underlined text.\n");
        var underlineProps = document.BeginUpdateCharacters(underlineRange);
        underlineProps.Underline = UnderlineType.Single;
        document.EndUpdateCharacters(underlineProps);

        return documentServer;
    }

    private async Task<RichEditDocumentServer> GenerateTableWithImages()
    {
        var documentServer = new RichEditDocumentServer();
        var document = documentServer.Document;

        // Create a table
        var table = document.Tables.Create(document.Range.End, 2, 2);

        // Insert images into the first row
        for (var colIndex = 0; colIndex < 2; colIndex++)
        {
            var cell = table[0, colIndex];
            var cellSubDoc = cell.Range.BeginUpdateDocument();

            // Load an image from a file or resource
            await using (var imageStream = await DownloadImageStreamAsync())
            {
                // Insert the image 
                document.Images.Insert(table.Rows[colIndex].Cells[1].Range.Start, DocumentImageSource.FromStream(imageStream));
                // Optionally, scale the image
                //image.ScaleX = 50; // Scale to 50% of the original size
                //image.ScaleY = 50;


                document.InsertText(table.Rows[colIndex].Cells[0].Range.Start, $"Caption {colIndex + 1}");
            }

            cell.Range.EndUpdateDocument(cellSubDoc);
        }

        return documentServer;
    }

    // Helper method to get image stream
    private Stream GetImageStream()
    {
        // Replace with your own method of obtaining an image stream
        // For example, reading from a file:
        // return System.IO.File.OpenRead("path_to_image.jpg");

        // For this example, create a simple bitmap in memory
        var bitmap = new Bitmap(100, 100);
        using (var g = Graphics.FromImage(bitmap))
        {
            g.Clear(Color.Red);
        }
        var ms = new MemoryStream();
        bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
        ms.Position = 0;
        return ms;
    }

    public static async Task<Stream> DownloadImageStreamAsync()
    {
        const string url = "https://picsum.photos/100/100";
        using var client = new HttpClient();

        // Download the image data as a byte array
        var imageData = await client.GetByteArrayAsync(url);

        // Create a MemoryStream from the image data
        var imageStream = new MemoryStream(imageData);

        // Set the stream position to the beginning
        imageStream.Position = 0;

        return imageStream;
    }

    private RichEditDocumentServer GenerateComplexTable()
    {
        var documentServer = new RichEditDocumentServer();

        var document = documentServer.Document;// as DevExpress.XtraRichEdit.API.Native.Implementation.NativeDocument;

        // Create a table with 4 rows and 4 columns
        var table = document.Tables.Create(document.Range.End, 4, 4);

        // Set table style
        table.TableLayout = TableLayoutType.Fixed;
        table.TableAlignment = TableRowAlignment.Center;

        // Apply formatting to the table cells
        foreach (var row in table.Rows)
        {
            foreach (var cell in row.Cells)
            {
                // Set cell background color
                cell.BackgroundColor = Color.LightGray;

                // Access the cell's document
                var cellSubDoc = cell.Range.BeginUpdateDocument();

                // Insert formatted text
                var cp = cellSubDoc.BeginUpdateCharacters(cellSubDoc.Range);
                cp.Bold = true;
                cp.ForeColor = Color.Blue;
                cellSubDoc.InsertText(cellSubDoc.CreatePosition(cellSubDoc.Range.End.ToInt()), "Formatted Text");
                cellSubDoc.EndUpdateCharacters(cp);

                cell.Range.EndUpdateDocument(cellSubDoc);
            }
        }

        // Merge some cells as an example
        table.MergeCells(table[0, 0], table[1, 1]);

        return documentServer;
    }
}
