using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2013.Word.Drawing;

/*  
 *  MasterPrinter is the printing class. It has two functions: Print() and FinalPrint().
 *  Print() is used to print last month's master file and the competing file.
 *  FinalPrint() is used to print FinalCollection and add the summary page, among other things.
 */
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
                currentRow++;
                ws.Cell(currentRow, 1).Value = "Total";
                int column = 2;
                foreach (string product in products)
                {
                    ws.Cell(currentRow, column).Value = product;
                    column++;
                }
                currentRow++;
                column = 2;
                int totalRow = currentRow;
                foreach (string product in products)
                {
                    ws.Cell(currentRow, column).Value = vendorDataSet.TotalSales(product);
                    column++;
                }
            }
            wb.SaveAs(vendorCollection.Name + ".xlsx");
        }

        public void FinalPrint(IVendorCollection vendorCollection, IVendorCollection masterCollection, IVendorCollection competingCollection, Dictionary<string, double> prices)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("GRID");
            ws.Cell("A1").Value = "Client";
            int currentRow = ws.FirstRowUsed().RowNumber() + 1;
            Dictionary<string, double> totalSpending = new Dictionary<string, double>();
            Dictionary<string, int> totalPurchases = new Dictionary<string, int>();
            foreach (VendorDataSet vendorDataSet in vendorCollection.GetVendorDataSets())
            {
                VendorDataSet masterDataSet = masterCollection.GetVendorDataSets()[0];
                VendorDataSet competingDataSet = competingCollection.GetVendorDataSets()[0];
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
                    bool competingExists = false;
                    CustomerVendorRecord masterCVR = customerVendorRecord;
                    CustomerVendorRecord competingCVR = customerVendorRecord;
                    if (masterDataSet.GetCustomers().Contains(customerVendorRecord.Customer))
                    {
                        masterCVR = masterDataSet.GetCustomerVendorRecords().Find(x => x.Customer.Equals(customerVendorRecord.Customer)) ?? throw new Exception();
                    }
                    if (competingDataSet.GetCustomers().Contains(customerVendorRecord.Customer))
                    {
                        competingCVR = competingDataSet.GetCustomerVendorRecords().Find(x => x.Customer.Equals(customerVendorRecord.Customer)) ?? throw new Exception();
                        competingExists = true;
                    }
                    ws.Cell(currentRow, 1).Value = customerVendorRecord.Customer;
                    int currentColumn = 2;
                    foreach (string product in products)
                    {
                        ws.Cell(currentRow, currentColumn).Value = customerVendorRecord.GetQuantity(product);
                        if (!competingExists)
                        {
                            if (customerVendorRecord.GetQuantity(product) == 0)
                            {
                                ws.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                                currentColumn++;
                                continue;
                            }
                            ws.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.Gold;
                            currentColumn++;
                            continue;
                        }
                        if (masterCVR.GetQuantity(product) > competingCVR.GetQuantity(product))
                        {
                            ws.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.Red;
                            currentColumn++;
                            continue;
                        }
                        if (masterCVR.GetQuantity(product) < competingCVR.GetQuantity(product))
                        {
                            ws.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.Green;
                            currentColumn++;
                            continue;
                        }
                        if (customerVendorRecord.GetQuantity(product) == 0)
                        {
                            ws.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.LightGoldenrodYellow;
                            currentColumn++;
                            continue;
                        }
                        ws.Cell(currentRow, currentColumn).Style.Fill.BackgroundColor = XLColor.Gold;
                        currentColumn++;
                    }
                    currentRow++;
                }
                currentRow++;
                ws.Cell(currentRow, 1).Value = "Total";
                int column = 2;
                foreach (string product in products)
                {
                    ws.Cell(currentRow, column).Value = product;
                    column++;
                }
                currentRow++;
                ws.Cell(currentRow, 1).Value = "Quantity";
                column = 2;
                foreach (string product in products)
                {
                    ws.Cell(currentRow, column).Value = vendorDataSet.TotalSales(product);
                    totalPurchases.Add(product, vendorDataSet.TotalSales(product));
                    column++;
                }
                currentRow++;
                int totalRow = currentRow;
                ws.Cell(currentRow, 1).Value = "Revenue";
                column = 2;
                foreach (string product in products)
                {
                    ws.Cell(currentRow, column).Value = vendorDataSet.TotalSales(product) * prices[product];
                    totalSpending.Add(product, vendorDataSet.TotalSales(product) * prices[product]);
                    column++;
                }
            }
            var filePaths = Directory.GetFiles("input", "Product Billing*.xlsx");
            XLWorkbook master = new XLWorkbook();
            try
            {
                master = new XLWorkbook(filePaths.First());
            }
            catch (Exception e) { Console.WriteLine(e); }
            master.Worksheet("Summary").CopyTo(wb, "Summary");
            master.Dispose();
            var summary = wb.Worksheet("Summary");
            var sumColumn = summary.Column(1);
            while (!summary.Cell(2, sumColumn.ColumnNumber() + 2).IsEmpty() || !summary.Cell(2, sumColumn.ColumnNumber() + 3).IsEmpty())
            {
                sumColumn = sumColumn.ColumnRight();
            }
            sumColumn.InsertColumnsBefore(1);
            summary.Cell(2, sumColumn.ColumnNumber()).Value = "Current Month";
            UnderVendors combiner = new UnderVendors();
            Dictionary<string, double> combinedRevenues = combiner.combineRevenue(totalSpending);
            wb.SaveAs(vendorCollection.Name + ".xlsx");
        }
    }
}