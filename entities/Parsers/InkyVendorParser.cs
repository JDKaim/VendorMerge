using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class InkyVendorParser : ExcelVendorParser
    {

        public override string Name => "Inky Product Billing";

        public InkyVendorParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName, worksheetName)
        {
        }
        protected override VendorParserResults ParseInternal(IXLWorksheet ws, IVendorCollection dataStore, IXLWorksheet renamer)
        {
            int recordsParsed = 0;
            var firstRowUsed = ws.Row(131);
            var lastColumnUsed = ws.LastColumnUsed();
            var lastRowUsed = ws.LastRowUsed();
            var categoryRow = firstRowUsed.RowBelow();
            while (!categoryRow.Cell(2).IsEmpty())
            {
                var categoryColumn = ws.Column(7);
                string customer = ws.Cell(categoryRow.RowNumber(), 2).GetString();
                while (categoryRow.RowNumber() < lastRowUsed.RowNumber() + 1 && categoryColumn.ColumnNumber() < lastColumnUsed.ColumnNumber() + 1)
                {
                    string vendor = "Vendor";
                    string product = "Inky";
                    if (categoryColumn.ColumnNumber() == lastColumnUsed.ColumnNumber())
                    {
                        product = "Inky Encryption";
                    }
                    if (!string.IsNullOrWhiteSpace(ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString()))
                    {
                        int quantity = 0;
                        if (ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString() != "-")
                        {
                            quantity = int.Parse(ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString());
                        }
                        dataStore.AddCustomerRecordQuantity(vendor, customer, product, quantity);
                    }
                    recordsParsed++;
                    categoryColumn = categoryColumn.ColumnRight();
                }
                categoryRow = categoryRow.RowBelow();
            }
            return VendorParserResults.CreateSuccess(recordsParsed);
        }
    }
}