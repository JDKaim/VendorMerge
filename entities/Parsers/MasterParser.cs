using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class MasterParser : ExcelVendorParser
    {

        public override string Name => "Master Product Billing";

        private string WorksheetName;

        public MasterParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName)
        {
            this.WorksheetName = worksheetName;
        }
        protected override VendorParserResults ParseInternal(XLWorkbook wb, IVendorCollection dataStore, IXLWorksheet renamer)
        {
            IXLWorksheet ws;
            try
            {
                ws = wb.Worksheet(this.WorksheetName);
            }
            catch (Exception e)
            {
                return VendorParserResults.CreateError($"An error occurred while loading the file for '{this.Name}': {e.Message}");
            }
            int recordsParsed = 0;
            var firstRowUsed = ws.FirstRowUsed();
            var firstColumnUsed = ws.FirstColumnUsed();
            var lastColumnUsed = ws.LastColumnUsed();
            var lastRowUsed = ws.LastRowUsed();
            var categoryRow = firstRowUsed.RowBelow();
            while (!categoryRow.Cell(1).IsEmpty())
            {
                var categoryColumn = firstColumnUsed.ColumnRight();
                string customer = ws.Cell(categoryRow.RowNumber(), 1).GetString();
                while (categoryRow.RowNumber() < lastRowUsed.RowNumber() + 1 && categoryColumn.ColumnNumber() < lastColumnUsed.ColumnNumber() + 1)
                {
                    string vendor = "Vendor";
                    string product = ws.Cell(1, categoryColumn.ColumnNumber()).GetString();
                    int quantity = 0;
                    if (ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString() != "" && ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString() != " ")
                    {
                        quantity = int.Parse(ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString());
                    }
                    dataStore.AddCustomerRecordQuantity(vendor, customer, product, quantity);
                    recordsParsed++;
                    categoryColumn = categoryColumn.ColumnRight();
                }
                categoryRow = categoryRow.RowBelow();
            }
            return VendorParserResults.CreateSuccess(recordsParsed);
        }
    }
}