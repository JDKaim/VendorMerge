using VendorMerge;
using VendorMerge.Parsers;
using VendorMerge.Printers;
using ClosedXML.Excel;

Dictionary<string, string> renamers = RenamingParser.ParseRename("input", "Renaming.xlsx", "GRID");
VendorCollection dataStore = new VendorCollection("Master Sheet", renamers);
VendorCollection competingDataStore = new VendorCollection("Competing Sheet", renamers);
Dictionary<string, double> prices = new Dictionary<string, double>();

Dictionary<string, DocumentName> documentNames = DocumentNamesParser.ParseDocumentNames();

MasterParser masterParser = new MasterParser(documentNames["Master"].FileLocation, documentNames["Master"].SpreadsheetName, documentNames["Master"].WorksheetName);

List<IVendorParser> newVendorParsers =
    new List<IVendorParser>()
    {
        new ProwrkVendorParser(documentNames["Prowrk"].FileLocation, documentNames["Prowrk"].SpreadsheetName, documentNames["Prowrk"].WorksheetName),
        new PronetVendorParser(documentNames["Pronet"].FileLocation, documentNames["Pronet"].SpreadsheetName, documentNames["Pronet"].WorksheetName),
        new ProservVendorParser(documentNames["Proserv"].FileLocation, documentNames["Proserv"].SpreadsheetName, documentNames["Proserv"].WorksheetName),
        new BitdefenderVendorParser(documentNames["Bitdefender"].FileLocation, documentNames["Bitdefender"].SpreadsheetName, documentNames["Bitdefender"].WorksheetName),
        new BluevaultVendorParser(documentNames["Bluevault"].FileLocation, documentNames["Bluevault"].SpreadsheetName, documentNames["Bluevault"].WorksheetName),
        new MyglueVendorParser(documentNames["Myglue"].FileLocation, documentNames["Myglue"].SpreadsheetName, documentNames["Myglue"].WorksheetName),
        new Kb4VendorParser(documentNames["Kb4"].FileLocation, documentNames["Kb4"].SpreadsheetName, documentNames["Kb4"].WorksheetName),
        new S1completeVendorParser(documentNames["S1complete"].FileLocation, documentNames["S1complete"].SpreadsheetName, documentNames["S1complete"].WorksheetName),
        new VhostproesxVendorParser(documentNames["Vhostproesx"].FileLocation, documentNames["Vhostproesx"].SpreadsheetName, documentNames["Vhostproesx"].WorksheetName),
        new VhostprohvVendorParser(documentNames["Vhostprohv"].FileLocation, documentNames["Vhostprohv"].SpreadsheetName, documentNames["Vhostprohv"].WorksheetName),
        // new S1controlVendorParser(documentNames["S1control"].FileLocation, documentNames["S1control"].SpreadsheetName, documentNames["S1control"].WorksheetName),
        new VeeamVendorParser(documentNames["Veeam"].FileLocation, documentNames["Veeam"].SpreadsheetName, documentNames["Veeam"].WorksheetName),
        // new InkyVendorParser(documentNames["Inky"].FileLocation, documentNames["Inky"].SpreadsheetName, documentNames["Inky"].WorksheetName)
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
while (!pricesheet.Cell(currentRow.RowNumber(), 1).IsEmpty())
{
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
printer.FinalPrint(consolidator.Consolidate(dataStore, competingDataStore, renamers), dataStore, competingDataStore, prices, documentNames["Master"]);

Console.WriteLine("Data successfully printed.");

// Read in existing master workbook.

// Write out new master workbook accounting for previous workbook data.