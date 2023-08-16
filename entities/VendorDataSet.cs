public class VendorDataSet
{

    private Dictionary<string, CustomerVendorRecord> _customerVendorRecords = new Dictionary<string, CustomerVendorRecord>();

    private Dictionary<string, bool> _products = new Dictionary<string, bool>();

    public readonly string Vendor;

    public VendorDataSet(string vendor)
    {
        this.Vendor = vendor;
    }

    public void AddCustomerRecordQuantity(string customer, string product, int quantity)
    {
        if (!_products.ContainsKey(product))
        {
            _products.Add(product, true);
        }
        if (!this._customerVendorRecords.ContainsKey(customer))
        {
            this._customerVendorRecords.Add(customer, new CustomerVendorRecord(this.Vendor, customer));
        }
        this._customerVendorRecords[customer].AddQuantity(product, quantity);
    }

    public List<CustomerVendorRecord> GetCustomerVendorRecords()
    {
        return this._customerVendorRecords.Values.ToList();
    }

    public List<string> GetProducts() {
        return this._products.Keys.ToList();
    }
}