using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class VeeamVendorParser : ExcelVendorParser
    {

        public override string Name => "Veeam Product Billing";

        public VeeamVendorParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName, worksheetName)
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
                var categoryColumn = firstColumnUsed.ColumnRight().ColumnRight().ColumnRight();
                var productColumn = categoryColumn.ColumnLeft();
                string customer = ws.Cell(categoryRow.RowNumber(), 1).GetString();
                string vendor = "Vendor";
                string product = "Veeam " + categoryRow.Cell(productColumn.ColumnNumber()).Value.ToString();
                if (product == "Veeam Standard Server ") {
                    product = "Veeam Standard Server";
                }
                int quantity = 0;
                if (!string.IsNullOrWhiteSpace(ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString()))
                {
                    quantity = int.Parse(ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString());
                    dataStore.AddCustomerRecordQuantity(vendor, customer, product, quantity);
                }
                recordsParsed++;
                categoryRow = categoryRow.RowBelow();
            }
            return VendorParserResults.CreateSuccess(recordsParsed);
        }
    }
}