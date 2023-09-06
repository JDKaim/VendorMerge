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
        protected override VendorParserResults ParseInternal(IXLWorksheet ws, IVendorCollection dataStore)
        {
            // int recordsParsed = 0;
            // var firstRowUsed = ws.FirstRowUsed();
            // var firstColumnUsed = ws.FirstColumnUsed();
            // var categoryRow = firstRowUsed.RowBelow();
            // while (!categoryRow.RowBelow().RowBelow().Cell(1).IsEmpty())
            // {
            //     var categoryColumn = firstColumnUsed.ColumnRight();
            //     string customer = ws.Cell(categoryRow.RowNumber(), 1).GetString();
            //     if (string.IsNullOrWhiteSpace(customer)) {
            //         categoryRow = categoryRow.RowBelow();
            //         continue;
            //     }
            //     string vendor = "Vendor";
            //     string product = "PROWRK";
            //     if (!string.IsNullOrWhiteSpace(ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString()))
            //     {
            //         int quantity = int.Parse(ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString());
            //         dataStore.AddCustomerRecordQuantity(vendor, customer, product, quantity);
            //     }
            //     recordsParsed++;
            //     categoryRow = categoryRow.RowBelow();
            // }
            // return VendorParserResults.CreateSuccess(recordsParsed);


            int recordsParsed = 0;
            var firstRowUsed = ws.FirstRowUsed();
            var categoryColumn = ws.FirstColumnUsed().ColumnRight();
            var categoryRow = ws.Row(88);
            string first = ws.Cell(categoryRow.RowNumber(), 1).GetString();

            while (!categoryRow.Cell(1).IsEmpty())
            {
                string customer = ws.Cell(categoryRow.RowNumber(), 1).GetString();
                if (string.IsNullOrWhiteSpace(customer)) {
                    categoryRow = categoryRow.RowBelow();
                    continue;
                }
                string vendor = "Vendor";
                string product = "PROWRK";
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