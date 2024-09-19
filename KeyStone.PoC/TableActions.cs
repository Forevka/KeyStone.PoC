using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;

namespace KeyStone.PoC;

static class TableActions
{
    public static void CreateTable(Document document)
    {
        #region #CreateTable
        // Insert new table.
        var tbl = document.Tables.Create(document.Range.Start, 1, 3, AutoFitBehaviorType.AutoFitToWindow);
        // Create a table header.
        document.InsertText(tbl[0, 0].Range.Start, "Name");
        document.InsertText(tbl[0, 1].Range.Start, "Size");
        document.InsertText(tbl[0, 2].Range.Start, "DateTime");
        // Insert table data.
        var dirinfo = new DirectoryInfo("C:\\");
        try
        {
            tbl.BeginUpdate();
            foreach (var fi in dirinfo.GetFiles())
            {
                var row = tbl.Rows.Append();
                var cell = row.FirstCell;
                var fileName = fi.Name;
                var fileLength = String.Format("{0:N0}", fi.Length);
                var fileLastTime = String.Format("{0:g}", fi.LastWriteTime);
                document.InsertSingleLineText(cell.Range.Start, fileName);
                document.InsertSingleLineText(cell.Next.Range.Start, fileLength);
                document.InsertSingleLineText(cell.Next.Next.Range.Start, fileLastTime);
            }
            // Center the table header.
            foreach (var p in document.Paragraphs.Get(tbl.FirstRow.Range))
            {
                p.Alignment = ParagraphAlignment.Center;
            }
        }
        finally
        {
            tbl.EndUpdate();
        }
        #endregion #CreateTable
    }

    public static void CreateFixedTable(Document document)
    {
        #region #CreateFixedTable
        var table = document.Tables.Create(document.Range.Start, 3, 4);

        table.TableAlignment = TableRowAlignment.Center;
        table.TableLayout = TableLayoutType.Fixed;
        table.PreferredWidthType = WidthType.Fixed;
        table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4f);

        table.Rows[1].HeightType = HeightType.Exact;
        table.Rows[1].Height = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.8f);

        table[1, 2].PreferredWidthType = WidthType.Fixed;
        table[1, 2].PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f);
        table.EndUpdate();

        #endregion #CreateFixedTable
    }
    public static void ChangeTableColor(Document document)
    {
        #region #ChangeTableColor
        // Create a table.
        var table = document.Tables.Create(document.Range.Start, 3, 5, AutoFitBehaviorType.AutoFitToWindow);
        table.BeginUpdate();
        // Provide the space between table cells.
        // The distance between cells will be 4 mm.
        document.Unit = DevExpress.Office.DocumentUnit.Millimeter;
        table.TableCellSpacing = 2;
        // Change the color of empty space between cells.
        table.TableBackgroundColor = Color.Violet;
        //Change cell background color.
        table.ForEachCell(new TableCellProcessorDelegate(TableHelper.ChangeCellColor));
        table.ForEachCell(new TableCellProcessorDelegate(TableHelper.ChangeCellBorderColor));
        table.EndUpdate();
        #endregion #ChangeTableColor

    }
    #region #@ChangeTableColor
    public class TableHelper
    {
        public static void ChangeCellColor(TableCell cell, int i, int j)
        {
            cell.BackgroundColor = Color.Yellow;
        }

        public static void ChangeCellBorderColor(TableCell cell, int i, int j)
        {
            cell.Borders.Bottom.LineColor = Color.Red;
            cell.Borders.Left.LineColor = Color.Red;
            cell.Borders.Right.LineColor = Color.Red;
            cell.Borders.Top.LineColor = Color.Red;
        }
    }
    #endregion #@ChangeTableColor
    public static void CreateAndApplyTableStyle(Document document)
    {
        #region #CreateAndApplyTableStyle
        document.BeginUpdate();
        // Create a new table style.
        var tStyleMain = document.TableStyles.CreateNew();
        // Specify style characteristics.
        tStyleMain.AllCaps = true;
        tStyleMain.FontName = "Segoe Condensed";
        tStyleMain.FontSize = 14;
        tStyleMain.Alignment = ParagraphAlignment.Center;
        tStyleMain.TableBorders.InsideHorizontalBorder.LineStyle = TableBorderLineStyle.Dotted;
        tStyleMain.TableBorders.InsideVerticalBorder.LineStyle = TableBorderLineStyle.Dotted;
        tStyleMain.TableBorders.Top.LineThickness = 1.5f;
        tStyleMain.TableBorders.Top.LineStyle = TableBorderLineStyle.Double;
        tStyleMain.TableBorders.Left.LineThickness = 1.5f;
        tStyleMain.TableBorders.Left.LineStyle = TableBorderLineStyle.Double;
        tStyleMain.TableBorders.Bottom.LineThickness = 1.5f;
        tStyleMain.TableBorders.Bottom.LineStyle = TableBorderLineStyle.Double;
        tStyleMain.TableBorders.Right.LineThickness = 1.5f;
        tStyleMain.TableBorders.Right.LineStyle = TableBorderLineStyle.Double;
        tStyleMain.CellBackgroundColor = Color.LightBlue;
        tStyleMain.TableLayout = TableLayoutType.Fixed;
        tStyleMain.Name = "MyTableStyle";
        //Add the style to the document.
        document.TableStyles.Add(tStyleMain);
        document.EndUpdate();
        document.BeginUpdate();
        // Create a table.
        var table = document.Tables.Create(document.Range.Start, 3, 3);
        table.TableLayout = TableLayoutType.Fixed;
        table.PreferredWidthType = WidthType.Fixed;
        table.PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(4.5f);
        table[1, 1].PreferredWidthType = WidthType.Fixed;
        table[1, 1].PreferredWidth = DevExpress.Office.Utils.Units.InchesToDocumentsF(1.5f);
        // Apply a previously defined style.
        table.Style = tStyleMain;
        document.EndUpdate();

        document.InsertText(table[1, 1].Range.Start, "STYLED");
        #endregion #CreateAndApplyTableStyle
    }

    public static void ChangeColumnAppearance(Document document)
    {
        #region #ChangeColumnAppearance
        var table = document.Tables.Create(document.Range.Start, 3, 10);
        table.BeginUpdate();
        //Change cell background color and vertical alignment in the third column.
        table.ForEachRow(new TableRowProcessorDelegate(ChangeColumnAppearanceHelper.ChangeColumnColor));
        table.EndUpdate();
        #endregion #ChangeColumnAppearance

    }
    #region #@ChangeColumnAppearance
    public class ChangeColumnAppearanceHelper
    {
        public static void ChangeColumnColor(TableRow row, int rowIndex)
        {
            row[2].BackgroundColor = Color.LightCyan;
            row[2].VerticalAlignment = TableCellVerticalAlignment.Center;
        }
    }
    #endregion #@ChangeColumnAppearance

    public static void UseTableCellProcessor(Document document)
    {
        #region #UseTableCellProcessor
        var table = document.Tables.Create(document.Range.Start, 8, 8);
        table.BeginUpdate();
        table.ForEachCell(new TableCellProcessorDelegate(UseTableCellProcessorHelper.MakeMultiplicationCell));
        table.EndUpdate();
        #endregion #UseTableCellProcessor
    }
    #region #@UseTableCellProcessor
    public class UseTableCellProcessorHelper
    {
        public static void MakeMultiplicationCell(TableCell cell, int i, int j)
        {
            var doc = cell.Range.BeginUpdateDocument();
            doc.InsertText(cell.Range.Start,
                String.Format("{0}*{1} = {2}", i + 2, j + 2, (i + 2) * (j + 2)));
            cell.Range.EndUpdateDocument(doc);
        }
    }
    #endregion #@UseTableCellProcessor

    public static void MergeCells(Document document)
    {
        #region #MergeCells
        var table = document.Tables.Create(document.Range.Start, 6, 8);
        table.BeginUpdate();
        table.MergeCells(table[2, 1], table[5, 1]);
        table.MergeCells(table[2, 3], table[2, 7]);
        table.EndUpdate();
        #endregion #MergeCells
    }
    public static void SplitCells(Document document)
    {
        #region #SplitCells
        var table = document.Tables.Create(document.Range.Start, 3, 3, AutoFitBehaviorType.FixedColumnWidth, 350);
        //split a cell to three: 
        table.Cell(2, 1).Split(1, 3);
        #endregion #SplitCells
    }
    public static void DeleteTableElements(Document document)
    {
        #region #DeleteTableElements
        var tbl = document.Tables.Create(document.Range.Start, 3, 3, AutoFitBehaviorType.AutoFitToWindow);
        tbl.BeginUpdate();
        //Delete a cell:
        tbl.Cell(1, 1).Delete();
        //Delete a row:
        tbl.Rows[2].Delete();
        tbl.EndUpdate();
        #endregion #DeleteTableElements
    }
    public static void DeleteTable(Document document)
    {
        #region #DeleteTable
        var tbl = document.Tables.Create(document.Range.Start, 3, 4);
        //To delete the table, uncomment the method below:
        //  document.Tables.Remove(tbl);
        #endregion #DeleteTable 
    }
    public static void WrapTextAroundTable(Document document)
    {
        #region #WrapTextAroundTable
        document.LoadDocument("Documents//Grimm.docx");

        var table = document.Tables.Create(document.Paragraphs[4].Range.Start, 3, 3, AutoFitBehaviorType.AutoFitToContents);

        table.BeginUpdate();
        table.TextWrappingType = TableTextWrappingType.Around;

        //Specify vertical alignment:
        table.RelativeVerticalPosition = TableRelativeVerticalPosition.Paragraph;
        table.VerticalAlignment = TableVerticalAlignment.None;
        table.OffsetYRelative = DevExpress.Office.Utils.Units.InchesToDocumentsF(2f);

        //Specify horizontal alignment:
        table.RelativeHorizontalPosition = TableRelativeHorizontalPosition.Margin;
        table.HorizontalAlignment = TableHorizontalAlignment.Center;

        //Set distance between the text and the table:
        table.MarginBottom = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
        table.MarginLeft = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
        table.MarginTop = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
        table.MarginRight = DevExpress.Office.Utils.Units.InchesToDocumentsF(0.3f);
        table.EndUpdate();
        #endregion #WrapTextAroundTable
    }
}