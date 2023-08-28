using ClosedXML.Excel;
public class UnderVendors
{

    private Dictionary<string, string> vendorProducts;

    public UnderVendors()
    {
        var vendorPaths = Directory.GetFiles("input", "Prices.xlsx");
        if (!vendorPaths.Any()) { throw new Exception($"Could not locate renaming file. Please insert it as \"prices.xlsx\"."); }
        if (vendorPaths.Length > 1) { throw new Exception($"Multiple renaming files located at \"prices.xlsx\". Please remove all but one."); }
        XLWorkbook vendorbook;
        try
        {
            vendorbook = new XLWorkbook(vendorPaths.First());
        }
        catch (Exception e)
        {
            throw new Exception($"An error occurred while loading the file for 'vendor.xlsx': {e.Message}");
        }
        IXLWorksheet vendorsheet;
        try
        {
            vendorsheet = vendorbook.Worksheet("Sheet1");
        }
        catch (Exception e)
        {
            throw new Exception($"An error occurred while loading the file for 'vendor.xlsx': {e.Message}");
        }
        var currentRow = vendorsheet.Row(1);
        while (!vendorsheet.Cell(currentRow.RowNumber(), 1).IsEmpty())
        {
            vendorProducts.Add(vendorsheet.Cell(currentRow.RowNumber(), 1).Value.ToString(), vendorsheet.Cell(currentRow.RowNumber(), 2).GetString());
            currentRow = currentRow.RowBelow();
        }
    }


    public Dictionary<string, double> combineRevenue(Dictionary<string, double> revenues)
    {
        List<string> products = revenues.Keys.ToList();
        Dictionary<string, double> combinedRevenues = new Dictionary<string, double>();
        foreach(string product in products) {
            string vendor = this.vendorProducts[product];
            if (combinedRevenues.ContainsKey(vendor)) {
                combinedRevenues[vendor] += revenues[product];
                continue;
            }
            combinedRevenues.Add(vendor, revenues[product]);
        }
        return combinedRevenues;
    }
}