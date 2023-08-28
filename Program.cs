using VendorMerge;
using VendorMerge.Parsers;
using VendorMerge.Printers;
using ClosedXML.Excel;

VendorCollection dataStore = new VendorCollection();
CompetingVendorCollection competingDataStore = new CompetingVendorCollection();
Dictionary<string, double> prices = new Dictionary<string, double>();

MasterParser masterParser = new MasterParser("input");

List<IVendorParser> newVendorParsers =
    new List<IVendorParser>()
    {
        new ProwrkVendorParser("input"),
        new PronetVendorParser("input"),
        new ProservVendorParser("input"),
        new BitdefenderVendorParser("input"),
        new BluevaultVendorParser("input"),
        new MyglueVendorParser("input"),
        new Kb4VendorParser("input"),
        new S1completeVendorParser("input"),
        new S1controlVendorParser("input"),
        new VeeamVendorParser("input"),
        new InkyVendorParser("input")
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
printer.FinalPrint(consolidator.Consolidate(dataStore, competingDataStore), dataStore, competingDataStore, prices);

Console.WriteLine("Data successfully printed.");

// Read in existing master workbook.

// Write out new master workbook accounting for previous workbook data.