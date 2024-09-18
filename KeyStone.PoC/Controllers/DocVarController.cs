using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit;
using Microsoft.AspNetCore.Mvc;
using System.Drawing;
using System.Linq;

namespace KeyStone.PoC.Controllers;

[ApiController]
[Route("api/[controller]")]
public class DocVarController : ControllerBase
{
    [HttpPost("CalculateDocumentVariable")]
    public IActionResult CalculateDocumentVariable([FromForm] string variableName, [FromForm] string? arguments)
    {
        /*
        // Create a new RichEditDocumentServer instance
        using RichEditDocumentServer documentServer = new RichEditDocumentServer();

        // Generate the content based on the variable name
        if (variableName.ToLower() == "table")
        {
            // Access the document
            Document document = documentServer.Document;

            // Insert a table at the start of the document
            Table table = document.Tables.Create(document.Range.Start, 3, 4);

            // Populate the table cells
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                TableRow row = table.Rows[rowIndex];
                for (int colIndex = 0; colIndex < row.Cells.Count; colIndex++)
                {
                    TableCell cell = row.Cells[colIndex];
                    SubDocument cellSubDoc = cell.Range.BeginUpdateDocument();
                    cellSubDoc.InsertText(cellSubDoc.CreatePosition(cellSubDoc.Range.End.ToInt()), $"{variableName} {rowIndex + 1},{colIndex + 1}");
                    cell.Range.EndUpdateDocument(cellSubDoc);
                }
            }

            // Get the RTF content of the document
            string rtfContent = documentServer.RtfText;//document.GetRtfText(document.);

            // Return the RTF content
            return Content(Base64Encode(rtfContent), "text/plain");
        }

        // Handle other variables or return empty content
        return Content(string.Empty, "text/plain");*/
        var rtfContent = string.Empty;

        switch (variableName)
        {
            case "ComplexTable":
                rtfContent = GenerateComplexTable();
                break;
            case "TableWithImages":
                rtfContent = GenerateTableWithImages();
                break;
            case "FormattedText":
                rtfContent = GenerateFormattedText();
                break;
            default:
                // Handle other variables or return empty content
                rtfContent = @"{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 Calibri;}}\f0\fs22 Unknown Variable}";
                break;
        }

        // Return the RTF content with appropriate content type
        return Content(Base64Encode(rtfContent), "text/plain");
    }

    private string GenerateFormattedText()
    {
        using var documentServer = new RichEditDocumentServer();

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

        // Get the RTF content
        return ToRtfBase64(documentServer);
    }

    private string GenerateTableWithImages()
    {
        using var documentServer = new RichEditDocumentServer();
        var document = documentServer.Document;

        // Create a table
        var table = document.Tables.Create(document.Range.End, 2, 2);

        // Insert images into the first row
        for (var colIndex = 0; colIndex < 2; colIndex++)
        {
            var cell = table[0, colIndex];
            var cellSubDoc = cell.Range.BeginUpdateDocument();

            // Load an image from a file or resource
            using (var imageStream = GetImageStream())
            {
                // Insert the image at the end of the cell's subdocument
                var image = cellSubDoc.Images.Insert(
                    cellSubDoc.CreatePosition(cellSubDoc.Range.End.ToInt()),
                    DocumentImageSource.FromStream(imageStream));

                // Optionally, scale the image
                image.ScaleX = 50; // Scale to 50% of the original size
                image.ScaleY = 50;
            }

            cell.Range.EndUpdateDocument(cellSubDoc);
        }

        // Insert text into the second row
        for (var colIndex = 0; colIndex < 2; colIndex++)
        {
            var cell = table[1, colIndex];
            var cellSubDoc = cell.Range.BeginUpdateDocument();
            cellSubDoc.InsertText(cellSubDoc.CreatePosition(cellSubDoc.Range.End.ToInt()), $"Caption {colIndex + 1}");
            cell.Range.EndUpdateDocument(cellSubDoc);
        }

        // **Save the entire document to RTF**
        return ToRtfBase64(documentServer);
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

    private string GenerateComplexTable()
    {
        using var documentServer = new RichEditDocumentServer();
        var document = documentServer.Document;

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

        // **Save the entire document to RTF**
        return ToRtfBase64(documentServer);
    }

    public string ToRtfBase64(RichEditDocumentServer documentServer)
    {
        using MemoryStream rtfStream = new MemoryStream();
        documentServer.SaveDocument(rtfStream, DocumentFormat.Rtf);
        rtfStream.Position = 0;
        using StreamReader reader = new StreamReader(rtfStream);
        var rtfContent = reader.ReadToEnd();

        return Base64Encode(rtfContent);
    }

    public static string Base64Encode(string plainText)
    {
        var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
        return System.Convert.ToBase64String(plainTextBytes);
    }
}
