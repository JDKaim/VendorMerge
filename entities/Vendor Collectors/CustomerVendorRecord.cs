public class CustomerVendorRecord
{

    private Dictionary<string, int> _quantities = new Dictionary<string, int>();

    public readonly string Vendor;
    public readonly string Customer;

    public CustomerVendorRecord(string vendor, string customer)
    {
        this.Vendor = vendor;
        this.Customer = customer;
    }

    public void AddQuantity(string product, int quantity)
    {
        if (quantity < 0) { throw new ArgumentOutOfRangeException("Cannot be a negative quantity."); }
        if (!this._quantities.ContainsKey(product))
        {
            this._quantities.Add(product, 0);
        }
        this._quantities[product] += quantity;
    }

    public void SubtractQuantity(string product, int quantity)
    {
        if (!this._quantities.ContainsKey(product))
        {
            throw new ArgumentOutOfRangeException("PROSERV hasn't been read and you are attempting to remove from it.");
        }
        if (this._quantities[product] < quantity)
        {
            throw new ArgumentOutOfRangeException("More VHOSTPRO-HV than PROSERV ):.");
        }
        this._quantities[product] -= quantity;
    }
    
    public int GetQuantity(string product)
    {
        if (!this._quantities.ContainsKey(product)) { return 0; }
        return this._quantities[product];
    }
}