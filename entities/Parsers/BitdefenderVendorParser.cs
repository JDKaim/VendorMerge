using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class BitdefenderVendorParser : ExcelVendorParser
    {

        public override string Name => "Bitdefender Product Billing";

        public BitdefenderVendorParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName, worksheetName)
        {
        }
        protected override VendorParserResults ParseInternal(IXLWorksheet ws, IVendorCollection dataStore)
        {
            int recordsParsed = 0;
            var firstRowUsed = ws.FirstRowUsed();
            var categoryRow = firstRowUsed.RowBelow();
            while (!categoryRow.Cell(1).IsEmpty())
            {
                string customer = ws.Cell(categoryRow.RowNumber(), 2).GetString();
                // customer = customer.Substring(1, customer.Length - 2);
                string vendor = "Vendor";
                string product = "Bitdefender";
                if (!string.IsNullOrWhiteSpace(ws.Cell(categoryRow.RowNumber(), 1).GetString()))
                {
                    int quantity = int.Parse(ws.Cell(categoryRow.RowNumber(), 1).GetString());
                    dataStore.AddCustomerRecordQuantity(vendor, customer, product, quantity);
                }
                recordsParsed++;
                categoryRow = categoryRow.RowBelow();
            }
            return VendorParserResults.CreateSuccess(recordsParsed);
        }
    }
}