using DocumentFormat.OpenXml.Wordprocessing;
using VendorMerge;

public class VendorCollection : IVendorCollection
{

    private Dictionary<string, VendorDataSet> _vendorDataSets = new Dictionary<string, VendorDataSet>();

    private readonly Dictionary<string, string> _Renamers;

    public string Name { get; private set; }

    public VendorCollection(string name, Dictionary<string, string> renamers)
    {
        this.Name = name;
        this._Renamers = renamers;
    }

    public void AddCustomerRecordQuantity(string vendor, string customer, string product, int quantity)
    {
        string realCustomer = "";
        if (_Renamers.ContainsKey(customer))
        {
            realCustomer = _Renamers[customer];
        }
        else
        {
            Console.WriteLine($"Customer '{customer}' does not exist. Please define in \"Renaming.xlsx\" or add to Master Sheet.");
        }
        if (!this._vendorDataSets.ContainsKey(vendor))
        {
            this._vendorDataSets.Add(vendor, new VendorDataSet(vendor));
        }
        this._vendorDataSets[vendor].AddCustomerRecordQuantity(realCustomer, product, quantity);
    }

    public List<VendorDataSet> GetVendorDataSets()
    {
        return this._vendorDataSets.Values.ToList();
    }

    public List<string> GetVendors()
    {
        return this._vendorDataSets.Keys.ToList();
    }
}