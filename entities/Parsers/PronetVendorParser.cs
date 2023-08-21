using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace VendorMerge.Parsers
{
    public class PronetVendorParser : ExcelVendorParser
    {

        public override string Name => "PRONET Product Billing";

        public PronetVendorParser(string inputDirectory) : base(inputDirectory, "Billing_AutomatePRONET_.xlsx")
        {

        }
        protected override VendorParserResults ParseInternal(XLWorkbook wb, IVendorCollection dataStore, IXLWorksheet renamer)
        {
            IXLWorksheet ws;
            try
            {
                ws = wb.Worksheet("in");
            }
            catch (Exception e)
            {
                return VendorParserResults.CreateError($"An error occurred while loading the file for '{this.Name}': {e.Message}");
            }
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
                if (ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnRight().ColumnNumber()).GetString() != "" && ws.Cell(categoryRow.RowNumber(), categoryColumn.ColumnRight().ColumnNumber()).GetString() != " ")
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
                    dataStore.AddCustomerRecordQuantity(vendor, customer, product, 1);
                }
                recordsParsed++;
                categoryRow = categoryRow.RowBelow();
            }
            return VendorParserResults.CreateSuccess(recordsParsed);
        }
    }
}