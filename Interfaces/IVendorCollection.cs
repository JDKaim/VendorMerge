namespace VendorMerge
{
    public interface IVendorCollection
    {
        string Name { get; }
        void AddCustomerRecordQuantity(string vendor, string customer, string product, int quantity);

        void SubtractCustomerRecordQuantity(string vendor, string customer, string product, int quantity);
        
        List<VendorDataSet> GetVendorDataSets();
        List<string> GetVendors();
    }
}
