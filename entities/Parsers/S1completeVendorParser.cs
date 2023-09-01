using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class S1completeVendorParser : ExcelVendorParser
    {

        public override string Name => "S1Complete Product Billing";

        public S1completeVendorParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName, worksheetName)
        {
        }
        protected override VendorParserResults ParseInternal(IXLWorksheet ws, IVendorCollection dataStore, IXLWorksheet renamer)
        {
            int recordsParsed = 0;
            var firstRowUsed = ws.FirstRowUsed();
            var firstColumnUsed = ws.FirstColumnUsed();
            var lastColumnUsed = ws.LastColumnUsed();
            var lastRowUsed = ws.LastRowUsed();
            var categoryRow = firstRowUsed.RowBelow();
            while (!categoryRow.Cell(1).IsEmpty())
            {
                var categoryColumn = firstColumnUsed;
                string customer = ws.Cell(categoryRow.RowNumber(), 2).GetString();
                customer = customer.Substring(2, customer.Length - 4);
                string vendor = "Vendor";
                string product = "SentinelOne Complete";
                int quantity = 0;
                if (!string.IsNullOrWhiteSpace(ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnRight().ColumnNumber()).GetString()))
                {
                    quantity = int.Parse(ws.Cell(categoryRow.RowNumber(), 1).GetString());
                    dataStore.AddCustomerRecordQuantity(vendor, customer, product, quantity);
                }
                recordsParsed++;
                categoryRow = categoryRow.RowBelow();
            }
            return VendorParserResults.CreateSuccess(recordsParsed);
        }
    }
}