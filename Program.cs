using VendorMerge;
using VendorMerge.Parsers;
using VendorMerge.Printers;

VendorCollection dataStore = new VendorCollection();
CompetingVendorCollection competingDataStore = new CompetingVendorCollection();

List<IVendorParser> vendorParsers =
    new List<IVendorParser>()
    {
        new MasterParser()
    };

List<IVendorParser> newVendorParsers =
    new List<IVendorParser>()
    {
        new FirstParser()
    };

foreach (IVendorParser vendorParser in vendorParsers)
{
    VendorParserResults results = vendorParser.Parse(dataStore);

    if (!results.Succeeded)
    {
        Console.Error.WriteLine("One or more errors occurred during vendor parsing: ");
        foreach (string error in results.Errors)
        {
            Console.Error.WriteLine($"- {error}");
        }
        return;
    }

    Console.WriteLine($"Parsed {results.RecordsParsed} from {vendorParser.Name}");
}

foreach (IVendorParser newVendorParser in newVendorParsers)
{
    VendorParserResults results = newVendorParser.Parse(competingDataStore);

    if (!results.Succeeded)
    {
        Console.Error.WriteLine("One or more errors occurred during vendor parsing: ");
        foreach (string error in results.Errors)
        {
            Console.Error.WriteLine($"- {error}");
        }
        return;
    }

    Console.WriteLine($"Parsed {results.RecordsParsed} from {newVendorParser.Name}");
}

Console.WriteLine($"All vendor files were successfully parsed");

MasterPrinter printer = new MasterPrinter();
printer.Print(dataStore);
printer.Print(competingDataStore);
Consolidator consolidator = new Consolidator();
printer.Print(consolidator.Consolidate(dataStore, competingDataStore));

Console.WriteLine("Data successfully printed.");

// Read in existing master workbook.

// Write out new master workbook accounting for previous workbook data.