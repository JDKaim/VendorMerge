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

        protected readonly string WorksheetName;

        public ExcelVendorParser(string inputDirectory, string filePattern, string worksheetName)
        {
            this.InputDirectory = inputDirectory;
            this.FilePattern = filePattern;
            this.WorksheetName = worksheetName;
        }

        abstract protected VendorParserResults ParseInternal(IXLWorksheet ws, IVendorCollection dataStore);

        public VendorParserResults Parse(IVendorCollection dataStore)
        {
            SheetParser sheetParser = new SheetParser(this.InputDirectory, this.FilePattern, this.WorksheetName);
            List<List<string>> data = sheetParser.GetTable();
            
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
            IXLWorksheet dataws;
            try
            {
                dataws = wb.Worksheet(this.WorksheetName);
            }
            catch (Exception e)
            {
                return VendorParserResults.CreateError($"An error occurred while loading the file for '{this.Name}': {e.Message}");
            }
            
            VendorParserResults vpr = this.ParseInternal(dataws, dataStore);
            wb.Dispose();
            return vpr;
        }
    }
}