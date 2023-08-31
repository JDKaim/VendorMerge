using VendorMerge;
using VendorMerge.Parsers;
using VendorMerge.Printers;
using ClosedXML.Excel;

VendorCollection dataStore = new VendorCollection();
CompetingVendorCollection competingDataStore = new CompetingVendorCollection();
Dictionary<string, double> prices = new Dictionary<string, double>();

Dictionary<string, DocumentName> documentNames = new Dictionary<string, DocumentName>();

var docPaths = Directory.GetFiles("input", "Document Names.xlsx");
if (!docPaths.Any()) { throw new Exception($"Could not locate renaming file. Please insert it as \"Document Names.xlsx\"."); }
if (docPaths.Length > 1) { throw new Exception($"Multiple renaming files located at \"Document Names.xlsx\". Please remove all but one."); }
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
while (!docsheet.Cell(docRow.RowNumber(), 1).IsEmpty()) {
    documentNames.Add(docsheet.Cell(docRow.RowNumber(), 1).Value.ToString(), new DocumentName(docsheet.Cell(docRow.RowNumber(), 2).Value.ToString(), docsheet.Cell(docRow.RowNumber(), 3).Value.ToString(), docsheet.Cell(docRow.RowNumber(), 4).Value.ToString()));
    docRow = docRow.RowBelow();
}


MasterParser masterParser = new MasterParser(documentNames["Master"].getFileLocation(), documentNames["Master"].getSpreadsheetName(), documentNames["Master"].getWorksheetName());

List<IVendorParser> newVendorParsers =
    new List<IVendorParser>()
    {
        new ProwrkVendorParser(documentNames["Prowrk"].getFileLocation(), documentNames["Prowrk"].getSpreadsheetName(), documentNames["Prowrk"].getWorksheetName()),
        new PronetVendorParser(documentNames["Pronet"].getFileLocation(), documentNames["Pronet"].getSpreadsheetName(), documentNames["Pronet"].getWorksheetName()),
        new ProservVendorParser(documentNames["Proserv"].getFileLocation(), documentNames["Proserv"].getSpreadsheetName(), documentNames["Proserv"].getWorksheetName()),
        new BitdefenderVendorParser(documentNames["Bitdefender"].getFileLocation(), documentNames["Bitdefender"].getSpreadsheetName(), documentNames["Bitdefender"].getWorksheetName()),
        new BluevaultVendorParser(documentNames["Bluevault"].getFileLocation(), documentNames["Bluevault"].getSpreadsheetName(), documentNames["Bluevault"].getWorksheetName()),
        new MyglueVendorParser(documentNames["Myglue"].getFileLocation(), documentNames["Myglue"].getSpreadsheetName(), documentNames["Myglue"].getWorksheetName()),
        new Kb4VendorParser(documentNames["Kb4"].getFileLocation(), documentNames["Kb4"].getSpreadsheetName(), documentNames["Kb4"].getWorksheetName()),
        new S1completeVendorParser(documentNames["S1complete"].getFileLocation(), documentNames["S1complete"].getSpreadsheetName(), documentNames["S1complete"].getWorksheetName()),
        new S1controlVendorParser(documentNames["S1control"].getFileLocation(), documentNames["S1control"].getSpreadsheetName(), documentNames["S1control"].getWorksheetName()),
        new VeeamVendorParser(documentNames["Veeam"].getFileLocation(), documentNames["Veeam"].getSpreadsheetName(), documentNames["Veeam"].getWorksheetName()),
        new InkyVendorParser(documentNames["Inky"].getFileLocation(), documentNames["Inky"].getSpreadsheetName(), documentNames["Inky"].getWorksheetName())
    };

VendorParserResults results = masterParser.Parse(dataStore);

var pricePaths = Directory.GetFiles("input", "Prices.xlsx");
if (!pricePaths.Any()) { throw new Exception($"Could not locate renaming file. Please insert it as \"prices.xlsx\"."); }
if (pricePaths.Length > 1) { throw new Exception($"Multiple renaming files located at \"prices.xlsx\". Please remove all but one."); }
XLWorkbook pricebook;
try
{
    pricebook = new XLWorkbook(pricePaths.First());
}
catch (Exception e)
{
    throw new Exception($"An error occurred while loading the file for 'prices.xlsx': {e.Message}");
}
IXLWorksheet pricesheet;
try
{
    pricesheet = pricebook.Worksheet("Sheet1");
}
catch (Exception e)
{
    throw new Exception($"An error occurred while loading the file for 'prices.xlsx': {e.Message}");
}

var currentRow = pricesheet.Row(1);
while (!pricesheet.Cell(currentRow.RowNumber(), 1).IsEmpty()) {
    prices.Add(pricesheet.Cell(currentRow.RowNumber(), 1).Value.ToString(), double.Parse(pricesheet.Cell(currentRow.RowNumber(), 2).GetString()));
    currentRow = currentRow.RowBelow();
}

Console.WriteLine($"Parsed {results.RecordsParsed} records from {masterParser.Name}");

foreach (IVendorParser newVendorParser in newVendorParsers)
{
    results = newVendorParser.Parse(competingDataStore);

    if (!results.Succeeded)
    {
        Console.Error.WriteLine("One or more errors occurred during vendor parsing: ");
        foreach (string error in results.Errors)
        {
            Console.Error.WriteLine($"- {error}");
        }
        return;
    }

    Console.WriteLine($"Parsed {results.RecordsParsed} records from {newVendorParser.Name}");
}

Console.WriteLine($"All vendor files were successfully parsed");

MasterPrinter printer = new MasterPrinter();
printer.Print(dataStore);
printer.Print(competingDataStore);
Consolidator consolidator = new Consolidator();
printer.FinalPrint(consolidator.Consolidate(dataStore, competingDataStore), dataStore, competingDataStore, prices, documentNames["Master"]);

Console.WriteLine("Data successfully printed.");

// Read in existing master workbook.

// Write out new master workbook accounting for previous workbook data.