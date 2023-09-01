using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using VendorMerge;

public class SheetParser : ITableDataSource
{
    private readonly string InputDirectory;
    private readonly string WorkbookName;
    private readonly string WorksheetName;
    public SheetParser(string inputDirectory, string workbookName, string worksheetName)
    {
        this.InputDirectory = inputDirectory;
        this.WorkbookName = workbookName;
        this.WorksheetName = worksheetName;
    }

    public List<List<string>> GetTable()
    {
        var filePaths = Directory.GetFiles(this.InputDirectory, this.WorkbookName);
        if (!filePaths.Any()) { throw new Exception($"Could not locate a file with name '{this.WorkbookName}' in the '{this.InputDirectory}' folder."); }
        if (filePaths.Length > 1) { throw new Exception($"Multiple files located with name '{this.WorkbookName}' in the '{this.InputDirectory}' folder. Please remove all but one."); }

        XLWorkbook wb;
        try
        {
            wb = new XLWorkbook(filePaths.First());
        }
        catch (Exception e)
        {
            throw new Exception($"An error occurred while loading the '{this.WorkbookName}' file: {e.Message}");
        }
        IXLWorksheet ws;
        try
        {
            if (string.IsNullOrWhiteSpace(this.WorksheetName))
            {
                ws = wb.Worksheets.First();
            }
            else
            {
                ws = wb.Worksheet(this.WorksheetName);
            }
        }
        catch (Exception e)
        {
            throw new Exception($"An error occurred while loading the '{this.WorkbookName}' file: {e.Message}");
        }

        List<List<string>> data = new List<List<string>>();


        var row = ws.FirstRowUsed();
        var column = ws.FirstColumnUsed();
        int firstColumn = column.ColumnNumber();
        while (!ws.Cell(row.RowNumber(), firstColumn).IsEmpty())
        {
            List<string> currentColumn = new List<string>();
            while (!ws.Cell(row.RowNumber(), column.ColumnNumber()).IsEmpty())
            {
                currentColumn.Add(ws.Cell(row.RowNumber(), column.ColumnNumber()).Value.ToString());
                column = column.ColumnRight();
            }
            data.Add(currentColumn);
            row = row.RowBelow();
            column = ws.Column(firstColumn);
        }
        wb.Dispose();

        return data;
    }
}