using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class VhostprohvVendorParser : ExcelVendorParser
    {

        public override string Name => "VHOSTPRO-HV Product Billing";

        public VhostprohvVendorParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName, worksheetName)
        {
        }
        protected override VendorParserResults ParseInternal(IXLWorksheet ws, IVendorCollection dataStore)
        {
            int recordsParsed = 0;
            var firstRowUsed = ws.FirstRowUsed();
            var categoryColumn = ws.FirstColumnUsed().ColumnRight();
            var categoryRow = ws.Row(3);
            while (!categoryRow.Cell(1).IsEmpty())
            {
                string customer = ws.Cell(categoryRow.RowNumber(), 1).GetString();
                string vendor = "Vendor";
                string product = "VHOSTPRO";
                if (!string.IsNullOrWhiteSpace(ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString()))
                {
                    dataStore.AddCustomerRecordQuantity(vendor, customer, product, 1);
                    dataStore.SubtractCustomerRecordQuantity(vendor, customer, "PROSERV", 1);
                }
                recordsParsed++;
                categoryRow = categoryRow.RowBelow();
            }
            return VendorParserResults.CreateSuccess(recordsParsed);
        }
    }
}