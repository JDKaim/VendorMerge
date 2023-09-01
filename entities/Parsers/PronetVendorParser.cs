using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class PronetVendorParser : ExcelVendorParser
    {

        public override string Name => "PRONET Product Billing";

        public PronetVendorParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName, worksheetName)
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
                var categoryColumn = firstColumnUsed.ColumnRight();
                string customer = ws.Cell(categoryRow.RowNumber(), 1).GetString();
                string vendor = "Vendor";
                string product = "PRONET";
                if (!string.IsNullOrWhiteSpace(ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnRight().ColumnNumber()).GetString()))
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