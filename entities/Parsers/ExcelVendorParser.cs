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

        abstract protected VendorParserResults ParseInternal(IXLWorksheet ws, IVendorCollection dataStore, IXLWorksheet renamer);

        public VendorParserResults Parse(IVendorCollection dataStore)
        {
            var filePaths = Directory.GetFiles(this.InputDirectory, this.FilePattern);
            var renamerPaths = Directory.GetFiles(this.InputDirectory, "Renaming.xlsx");
            if (!filePaths.Any()) { return VendorParserResults.CreateError($"Could not locate file for '{this.Name}'"); }
            if (filePaths.Length > 1) { return VendorParserResults.CreateError($"Multiple files located for '{this.Name}'. Please remove all but one."); }
            if (!renamerPaths.Any()) { return VendorParserResults.CreateError($"Could not locate renaming file. Please insert it as \"renaming.xlsx\"."); }
            if (renamerPaths.Length > 1) { return VendorParserResults.CreateError($"Multiple renaming files located at \"renaming.xlsx\". Please remove all but one."); }
            XLWorkbook wb;
            XLWorkbook renamer;
            try
            {
                wb = new XLWorkbook(filePaths.First());
            }
            catch (Exception e)
            {
                return VendorParserResults.CreateError($"An error occurred while loading the file for '{this.Name}': {e.Message}");
            }
            try
            {
                renamer = new XLWorkbook(renamerPaths.First());
            }
            catch (Exception e)
            {
                return VendorParserResults.CreateError($"An error occurred while loading the file for renaming: {e.Message}");
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
            IXLWorksheet ws;
            try
            {
                ws = renamer.Worksheet("GRID");
            }
            catch (Exception e)
            {
                return VendorParserResults.CreateError($"An error occurred while loading the file for '{this.Name}': {e.Message}");
            }
            
            VendorParserResults vpr = this.ParseInternal(dataws, dataStore, ws);
            wb.Dispose();
            return vpr;
        }
    }
}