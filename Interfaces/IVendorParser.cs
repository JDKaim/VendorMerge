namespace VendorMerge
{
    public interface IVendorParser
    {
        string Name { get; }
        VendorParserResults Parse(IVendorCollection dataStore);
    }
}
