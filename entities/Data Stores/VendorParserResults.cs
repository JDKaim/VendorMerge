namespace VendorMerge
{
    public class VendorParserResults
    {
        public int RecordsParsed;
        public readonly List<string> Errors;

        public bool Succeeded => !this.Errors.Any();

        public VendorParserResults()
        {
            this.Errors = new List<string>();
        }

        public static VendorParserResults CreateSuccess(int recordsParsed)
        {
            VendorParserResults results = new VendorParserResults
            {
                RecordsParsed = recordsParsed
            };
            return results;
        }

        public static VendorParserResults CreateError(string message)
        {
            VendorParserResults results = new VendorParserResults();
            results.Errors.Add(message);
            return results;
        }
    }
}
