using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class InkyVendorParser : ExcelVendorParser
    {

        public override string Name => "Inky Product Billing";

        private string WorksheetName;

        public InkyVendorParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName)
        {
            this.WorksheetName = worksheetName;
        }
        protected override VendorParserResults ParseInternal(XLWorkbook wb, IVendorCollection dataStore, IXLWorksheet renamer)
        {
            IXLWorksheet ws;
            try
            {
                ws = wb.Worksheet(this.WorksheetName);
            }
            catch (Exception e)
            {
                return VendorParserResults.CreateError($"An error occurred while loading the file for '{this.Name}': {e.Message}");
            }
            int recordsParsed = 0;
            var firstRowUsed = ws.Row(131);
            var firstColumnUsed = ws.FirstColumnUsed();
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
                    if (categoryColumn.ColumnNumber() == lastColumnUsed.ColumnNumber()) {
                        product = "Inky Encryption";
                    }
                    int quantity = 0;
                    if (ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString() != "" && ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString() != " ")
                    {
                        bool contained = false;
                        foreach (VendorDataSet vendorDS in dataStore.GetVendorDataSets()) {
                            if (vendorDS.GetCustomers().Contains(customer)) {
                                contained = true;
                                break;
                            }
                        }
                        if (!contained) {
                            bool fixer = false;
                            var firstRenamingRow = renamer.FirstRowUsed();
                            while (!firstRenamingRow.IsEmpty()) {
                                if (firstRenamingRow.Cell(1).Value.ToString() == customer) {
                                    customer = firstRenamingRow.Cell(2).Value.ToString();
                                    fixer = true;
                                    break;
                                }
                                firstRenamingRow = firstRenamingRow.RowBelow();
                            }
                            if (!fixer) {
                                return VendorParserResults.CreateError($"Customer '{customer}' does not exist. Please define in \"Renaming.xlsx\" or add to Master Sheet.");
                            }
                        }
                        if (ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnNumber()).GetString() == "-") {
                            quantity = 0;
                            continue;
                        }
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