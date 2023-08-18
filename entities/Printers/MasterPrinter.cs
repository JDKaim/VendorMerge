using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;

namespace VendorMerge.Printers
{
    public class MasterPrinter
    {
        public void Print(IVendorCollection vendorCollection)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("GRID");
            ws.Cell("A1").Value = "Client";
            int currentRow = ws.FirstRowUsed().RowNumber() + 1;
            foreach (VendorDataSet vendorDataSet in vendorCollection.GetVendorDataSets())
            {
                List<string> products = vendorDataSet.GetProducts();
                int topColumn = 2;
                foreach (string product in products)
                {
                    ws.Cell(1, topColumn).Value = product;
                    topColumn++;
                }
                List<CustomerVendorRecord> customerVendorRecords = vendorDataSet.GetCustomerVendorRecords();
                foreach (CustomerVendorRecord customerVendorRecord in customerVendorRecords)
                {
                    ws.Cell(currentRow, 1).Value = customerVendorRecord.Customer;
                    int currentColumn = 2;
                    foreach (string product in products)
                    {
                        ws.Cell(currentRow, currentColumn).Value = customerVendorRecord.GetQuantity(product);
                        currentColumn++;
                    }
                    currentRow++;
                }
            }
            wb.SaveAs(vendorCollection.Name + ".xlsx");
        }
    }
}