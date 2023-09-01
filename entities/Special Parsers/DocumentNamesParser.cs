using ClosedXML.Excel;

public class DocumentNamesParser {
    public DocumentNamesParser() {
        
    }

    public static Dictionary<string, DocumentName> ParseDocumentNames() {
        Dictionary<string, DocumentName> documentNames = new Dictionary<string, DocumentName>();
        var docPaths = Directory.GetFiles("input", "Document Names.xlsx");
        if (!docPaths.Any()) { throw new Exception($"Could not locate document naming file. Please insert it as \"Document Names.xlsx\"."); }
        if (docPaths.Length > 1) { throw new Exception($"Multiple document naming files located at \"Document Names.xlsx\". Please remove all but one."); }
        XLWorkbook docbook;
        try
        {
            docbook = new XLWorkbook(docPaths.First());
        }
        catch (Exception e)
        {
            throw new Exception($"An error occurred while loading the file for 'Document Names.xlsx': {e.Message}");
        }
        IXLWorksheet docsheet;
        try
        {
            docsheet = docbook.Worksheet("Sheet1");
        }
        catch (Exception e)
        {
            throw new Exception($"An error occurred while loading the file for 'Document Names.xlsx': {e.Message}");
        }

        var docRow = docsheet.Row(2);
        while (!docsheet.Cell(docRow.RowNumber(), 1).IsEmpty())
        {
            documentNames.Add(docsheet.Cell(docRow.RowNumber(), 1).Value.ToString(), new DocumentName(docsheet.Cell(docRow.RowNumber(), 2).Value.ToString(), docsheet.Cell(docRow.RowNumber(), 3).Value.ToString(), docsheet.Cell(docRow.RowNumber(), 4).Value.ToString()));
            docRow = docRow.RowBelow();
        }
        docbook.Dispose();
        return documentNames;
    }
}