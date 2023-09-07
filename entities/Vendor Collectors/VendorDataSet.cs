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

    public void SubtractCustomerRecordQuantity(string customer, string product, int quantity)
    {
        if (!_products.ContainsKey(product)) { throw new ArgumentOutOfRangeException(product + " does not exist."); }
        if (!this._customerVendorRecords.ContainsKey(customer)) { throw new ArgumentOutOfRangeException(customer + " does not exist."); }
        this._customerVendorRecords[customer].SubtractQuantity(product, quantity);
    }

    public List<CustomerVendorRecord> GetCustomerVendorRecords()
    {
        return this._customerVendorRecords.Values.ToList();
    }

    public List<string> GetCustomers()
    {
        return this._customerVendorRecords.Keys.ToList();
    }

    public List<string> GetProducts()
    {
        return this._products.Keys.ToList();
    }

    public int TotalSales(string product)
    {
        int total = 0;
        foreach (CustomerVendorRecord cvr in GetCustomerVendorRecords())
        {
            total += cvr.GetQuantity(product);
        }
        return total;
    }
}