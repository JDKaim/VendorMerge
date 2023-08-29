using System.Reflection.Metadata.Ecma335;
using System.Reflection.PortableExecutable;
using System.Runtime.CompilerServices;
using ClosedXML.Excel;
using VendorMerge;

public class UnderVendors
{

    private Dictionary<string, string> _vendorProducts;

    public UnderVendors()
    {
        var vendorPaths = Directory.GetFiles("input", "Vendors.xlsx");
        this._vendorProducts = new Dictionary<string, string>();
        if (!vendorPaths.Any()) { throw new Exception($"Could not locate renaming file. Please insert it as \"vendors.xlsx\"."); }
        if (vendorPaths.Length > 1) { throw new Exception($"Multiple renaming files located at \"vendors.xlsx\". Please remove all but one."); }
        XLWorkbook vendorbook;
        try
        {
            vendorbook = new XLWorkbook(vendorPaths.First());
        }
        catch (Exception e)
        {
            throw new Exception($"An error occurred while loading the file for 'vendors.xlsx': {e.Message}");
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
            this._vendorProducts.Add(vendorsheet.Cell(currentRow.RowNumber(), 1).Value.ToString(), vendorsheet.Cell(currentRow.RowNumber(), 2).GetString());
            currentRow = currentRow.RowBelow();
        }
    }


    public Dictionary<string, double> combineRevenue(Dictionary<string, double> revenues)
    {
        List<string> products = revenues.Keys.ToList();
        Dictionary<string, double> combinedRevenues = new Dictionary<string, double>();
        foreach (string product in products)
        {
            string vendor = this._vendorProducts[product];
            if (combinedRevenues.ContainsKey(vendor))
            {
                combinedRevenues[vendor] += revenues[product];
                continue;
            }
            combinedRevenues.Add(vendor, revenues[product]);
        }
        return combinedRevenues;
    }

    public Dictionary<string, Dictionary<string, int>> combineProducts(IVendorCollection vendors)
    {
        Dictionary<string, Dictionary<string, int>> _customerProducts = new Dictionary<string, Dictionary<string, int>>();
        foreach (VendorDataSet vds in vendors.GetVendorDataSets())
        {
            foreach (CustomerVendorRecord cvr in vds.GetCustomerVendorRecords())
            {
                Dictionary<string, int> _customerVendors = new Dictionary<string, int>();
                foreach (string product in vds.GetProducts())
                {
                    int amount = cvr.GetQuantity(product);
                    string actualVendor = this._vendorProducts[product];
                    if (_customerVendors.ContainsKey(actualVendor))
                    {
                        _customerVendors[actualVendor] += amount;
                        if (_customerVendors.ContainsKey(product))
                        {
                            _customerVendors[product] += amount;
                            continue;
                        }
                        _customerVendors.Add(product, amount);
                        continue;
                    }
                    _customerVendors.Add(actualVendor, amount);
                    if (_customerVendors.ContainsKey(product))
                    {
                        _customerVendors[product] += amount;
                        continue;
                    }
                    _customerVendors.Add(product, amount);
                }
                _customerProducts.Add(cvr.Customer, _customerVendors);
            }
        }
        return _customerProducts;
    }

    public Dictionary<string, int> customerUsage(Dictionary<string, Dictionary<string, int>> _customerProducts)
    {
        Dictionary<string, int> _customersUsing = new Dictionary<string, int>();
        string newVendor = "KnowBe4";
        _customersUsing.Add(newVendor, 0);
        foreach (string customer in _customerProducts.Keys.ToList())
        {
            foreach (string vendor in _customerProducts[customer].Keys.ToList())
            {
                if (_customerProducts[customer][vendor] > 0)
                {
                    if (vendor == "KnowBe4 Bulk" || vendor == "KnowBe4 Custom") {
                        _customersUsing["KnowBe4"] += 1;
                    }
                    if (_customersUsing.ContainsKey(vendor))
                    {
                        _customersUsing[vendor] += 1;
                        continue;
                    }
                    _customersUsing.Add(vendor, 1);
                }
            }
        }
        return _customersUsing;
    }
}