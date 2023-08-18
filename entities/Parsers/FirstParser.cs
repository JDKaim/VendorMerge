using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class FirstParser : IVendorParser
    {

        public string Name => "Master Product Billing";

        public VendorParserResults Parse(IVendorCollection dataStore)
        {
            var wb = new XLWorkbook("input\\Billing_AutomatePROWRK_.xlsx");
            var ws = wb.Worksheet("in");
            bool foundVendorFile = false;
            if (foundVendorFile) { return VendorParserResults.CreateError($"Could not locate vendor file for '{this.Name}'"); }

            bool fileLoadErrorOccurred = false;
            if (fileLoadErrorOccurred) { return VendorParserResults.CreateError($"An error occurred while loading the file for '{this.Name}': (Error like 'file could not be loaded as an Excel file')"); }

            int recordsParsed = 0;
            var firstRowUsed = ws.FirstRowUsed();
            var firstColumnUsed = ws.FirstColumnUsed();
            var lastColumnUsed = ws.LastColumnUsed();
            var lastRowUsed = ws.LastRowUsed();
            var categoryRow = firstRowUsed.RowBelow();
            while (!categoryRow.RowBelow().RowBelow().Cell(1).IsEmpty())
            {
                var categoryColumn = firstColumnUsed.ColumnRight();
                string customer = ws.Cell(categoryRow.RowNumber(), 1).GetString();
                while (categoryRow.RowNumber() < lastRowUsed.RowNumber() + 1 && categoryColumn.ColumnNumber() < lastColumnUsed.ColumnNumber() + 1)
                {
                    string vendor = "Vendor";
                    string product = "PROWRK";
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