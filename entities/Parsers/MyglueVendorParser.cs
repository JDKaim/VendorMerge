using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class MyglueVendorParser : ExcelVendorParser
    {

        public override string Name => "MyGlue Product Billing";

        public MyglueVendorParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName, worksheetName)
        {
        }
        protected override VendorParserResults ParseInternal(IXLWorksheet ws, IVendorCollection dataStore)
        {
            int recordsParsed = 0;
            var firstRowUsed = ws.FirstRowUsed();
            var categoryColumn = ws.FirstColumnUsed().ColumnRight();
            var categoryRow = firstRowUsed.RowBelow();
            while (!categoryRow.Cell(1).IsEmpty())
            {
                string customer = ws.Cell(categoryRow.RowNumber(), 2).GetString();
                string vendor = "Vendor";
                string product = "MyGlue";
                if (!string.IsNullOrWhiteSpace(ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString()))
                {
                    dataStore.AddCustomerRecordQuantity(vendor, customer, product, 1);
                }
                recordsParsed++;
                categoryRow = categoryRow.RowBelow();
            }
            return VendorParserResults.CreateSuccess(recordsParsed);
        }
    }
}