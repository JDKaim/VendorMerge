using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    abstract public class ExcelVendorParser : IVendorParser
    {

        abstract public string Name { get; }

        protected readonly string InputDirectory;

        protected readonly string FilePattern;

        public ExcelVendorParser(string inputDirectory, string filePattern)
        {
            this.InputDirectory = inputDirectory;
            this.FilePattern = filePattern;
        }

        abstract protected VendorParserResults ParseInternal(XLWorkbook wb, IVendorCollection dataStore);

        public VendorParserResults Parse(IVendorCollection dataStore)
        {
            var filePaths = Directory.GetFiles(this.InputDirectory, this.FilePattern);
            if (!filePaths.Any()) { return VendorParserResults.CreateError($"Could not locate file for '{this.Name}'"); }
            if (filePaths.Length > 1) { return VendorParserResults.CreateError($"Multiple files located for '{this.Name}'. Please remove all but one."); }
            XLWorkbook wb;
            try
            {
                wb = new XLWorkbook(filePaths.First());
            }
            catch (Exception e)
            {
                return VendorParserResults.CreateError($"An error occurred while loading the file for '{this.Name}': {e.Message}");
            }
            VendorParserResults vpr = this.ParseInternal(wb, dataStore);
            wb.Dispose();
            return vpr;
        }
    }
}