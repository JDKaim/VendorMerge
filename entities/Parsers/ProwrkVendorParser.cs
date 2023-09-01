using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class ProwrkVendorParser : ExcelVendorParser
    {

        public override string Name => "PROWRK Product Billing";


        public ProwrkVendorParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName, worksheetName)
        {
        }
        protected override VendorParserResults ParseInternal(IXLWorksheet ws, IVendorCollection dataStore, IXLWorksheet renamer)
        {
            int recordsParsed = 0;
            var firstRowUsed = ws.FirstRowUsed();
            var firstColumnUsed = ws.FirstColumnUsed();
            var categoryRow = firstRowUsed.RowBelow();
            while (!categoryRow.RowBelow().RowBelow().Cell(1).IsEmpty())
            {
                var categoryColumn = firstColumnUsed.ColumnRight();
                string customer = ws.Cell(categoryRow.RowNumber(), 1).GetString();

                string vendor = "Vendor";
                string product = "PROWRK";
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