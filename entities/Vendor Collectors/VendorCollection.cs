using DocumentFormat.OpenXml.Wordprocessing;
using VendorMerge;

public class VendorCollection : IVendorCollection
{

    private Dictionary<string, VendorDataSet> _vendorDataSets = new Dictionary<string, VendorDataSet>();

    public string Name { get; private set; }

    public VendorCollection(string name)
    {
        this.Name = name;
    }

    public void AddCustomerRecordQuantity(string vendor, string customer, string product, int quantity)
    {
        if (!this._vendorDataSets.ContainsKey(vendor))
        {
            this._vendorDataSets.Add(vendor, new VendorDataSet(vendor));
        }
        this._vendorDataSets[vendor].AddCustomerRecordQuantity(customer, product, quantity);
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