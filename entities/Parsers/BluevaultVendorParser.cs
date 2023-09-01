using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class BluevaultVendorParser : ExcelVendorParser
    {

        public override string Name => "BlueVault Product Billing";

        public BluevaultVendorParser(string inputDirectory, string spreadsheetName, string worksheetName) : base(inputDirectory, spreadsheetName, worksheetName)
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
                var categoryColumn = firstColumnUsed;
                string customer = ws.Cell(categoryRow.RowNumber(), 1).GetString();
                string vendor = "Vendor";
                string product = "BlueVault";
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
                }
                dataStore.AddCustomerRecordQuantity(vendor, customer, product, 1);
                recordsParsed++;
                categoryRow = categoryRow.RowBelow();
            }
            return VendorParserResults.CreateSuccess(recordsParsed);
        }
    }
}