public class VendorCollection
{

    private Dictionary<string, VendorDataSet> _vendorDataSets = new Dictionary<string, VendorDataSet>();

    public VendorCollection()
    {
    }

    public void AddCustomerRecordQuantity(string vendor, string customer, string product, int quantity)
    {
        if (!_vendors.ContainsKey(vendor))
        {
            _vendors.Add(vendor, true);
        }
        if (!this._vendorDataSets.ContainsKey(vendor))
        {
            this._vendorDataSets.Add(vendor);
        }
        this._vendorDataSets[vendor].AddCustomerRecordQuantity(customer, product, quantity);
    }

    public List<VendorCollection> GetVendorDataSets()
    {
        return this._vendorDataSets.Values.ToList();
    }

    public List<string> GetVendors()
    {
        return this._vendorDataSets.Keys.ToList();
    }
}