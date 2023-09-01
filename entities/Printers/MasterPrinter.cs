using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2013.Word.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;

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

        public void FinalPrint(IVendorCollection vendorCollection, IVendorCollection masterCollection, IVendorCollection competingCollection, Dictionary<string, double> prices, DocumentName masterSheet)
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
            var filePaths = Directory.GetFiles(masterSheet.FileLocation, masterSheet.SpreadsheetName);
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
            sumColumn = sumColumn.ColumnLeft();
            summary.Cell(2, sumColumn.ColumnNumber()).Value = "Current Month";
            UnderVendors combiner = new UnderVendors();
            Dictionary<string, double> combinedRevenues = combiner.combineRevenue(totalSpending);
            var sumRow = summary.Row(3);
            double totalSum = 0;
            while (sumRow.RowNumber() < 14)
            {
                double currentSum = combinedRevenues[summary.Cell(sumRow.RowNumber(), 1).Value.ToString()];
                totalSum += currentSum;
                summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber()).Value = currentSum;
                double monthlyDifference = currentSum - double.Parse(summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() - 1).Value.ToString());
                summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 2).Value = currentSum + double.Parse(summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 2).Value.ToString());
                if (monthlyDifference < 0)
                {
                    summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 1).Value = $"({monthlyDifference * -1})";
                    summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 1).Style.Font.FontColor = XLColor.Red;
                    sumRow = sumRow.RowBelow();
                    continue;
                }
                summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 1).Value = monthlyDifference;
                sumRow = sumRow.RowBelow();
            }
            sumRow = sumRow.RowBelow().RowBelow();
            summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber()).Value = totalSum;
            summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 2).Value = totalSum + double.Parse(summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 2).Value.ToString());
            double totalDifference = totalSum - double.Parse(summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() - 1).Value.ToString());
            if (totalDifference < 0)
            {
                summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 1).Value = $"({totalDifference * -1})";
                summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 1).Style.Font.FontColor = XLColor.Red;
                summary.Cell(19, sumColumn.ColumnNumber() + 2).Style.Fill.BackgroundColor = XLColor.CherryBlossomPink;
                summary.Cell(19, sumColumn.ColumnNumber() + 2).Style.Font.FontColor = XLColor.Red;
            }
            else
            {
                summary.Cell(19, sumColumn.ColumnNumber() + 2).Style.Fill.BackgroundColor = XLColor.LightGreen;
                summary.Cell(19, sumColumn.ColumnNumber() + 2).Style.Font.FontColor = XLColor.Green;
                summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 1).Value = totalDifference;
            }
            sumRow = sumRow.RowBelow();
            double lastMonth = double.Parse(summary.Cell(sumRow.RowNumber() - 1, sumColumn.ColumnNumber() - 1).Value.ToString());
            summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber()).Value = (totalSum - lastMonth) / lastMonth;
            summary.Cell(19, sumColumn.ColumnNumber() + 2).Value = (totalSum - lastMonth) / lastMonth;


            Dictionary<string, Dictionary<string, int>> _vendorProducts = combiner.combineProducts(vendorCollection);
            sumRow = sumRow.RowBelow().RowBelow().RowBelow().RowBelow().RowBelow().RowBelow();
            Dictionary<string, int> _customersUsing = combiner.customerUsage(_vendorProducts);
            sumColumn = summary.Column(sumColumn.ColumnNumber() - 8);
            List<string> characters = new List<string>{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
            int sumColumnNumber = sumColumn.ColumnNumber();
            string sumLetters = "";
            while (sumColumnNumber > 0) {
                sumLetters = characters[(sumColumnNumber % 26) - 1] + sumLetters;
                sumColumnNumber /= 26;
            }
            summary.Range($"{sumLetters}{sumRow.RowNumber() - 2}:{sumLetters}{39}").InsertColumnsBefore(1);
            summary.Cell(sumRow.RowNumber() - 1, sumColumn.ColumnNumber()).Value = "Current Month";
            while (sumRow.RowNumber() < 40)
            {
                string currentVendor = summary.Cell(sumRow.RowNumber(), 1).Value.ToString();
                int currentNum = 0;
                if (_customersUsing.ContainsKey(currentVendor))
                {
                    currentNum = _customersUsing[currentVendor];
                }
                summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber()).Value = currentNum;
                summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 1).Value = currentNum / double.Parse(summary.Cell(21, sumColumn.ColumnNumber() + 1).Value.ToString());
                summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() + 2).Value = currentNum - int.Parse(summary.Cell(sumRow.RowNumber(), sumColumn.ColumnNumber() - 1).Value.ToString());
                sumRow = sumRow.RowBelow();
            }
            
            wb.SaveAs(vendorCollection.Name + ".xlsx");
        }
    }
}