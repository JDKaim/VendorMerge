using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class VhostproesxVendorParser : ExcelVendorParser
    {

        public override string Name => "VHOSTPRO-ESX Product Billing";

        public VhostproesxVendorParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName, worksheetName)
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
                string customer = ws.Cell(categoryRow.RowNumber(), 1).GetString();
                string vendor = "Vendor";
                string product = "VHOSTPRO";
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