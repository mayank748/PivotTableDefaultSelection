using OfficeOpenXml;

namespace DefaultValuSelect
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            if (!stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                CreatePivotTable();
                _logger.LogInformation("File created EpPlusDefaultSelect.xlsx", DateTimeOffset.Now);
            }
        }

        public void CreatePivotTable()
        {
            // Initialize the EPPlus package
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Create a new Excel package
            using (ExcelPackage package = new ExcelPackage())
            {
                // Add a worksheet for raw data
                ExcelWorksheet dataSheet = package.Workbook.Worksheets.Add("RawData");

                // Add data to the sheet (replace this with your data)
                dataSheet.Cells["A1"].Value = "Fund";
                dataSheet.Cells["B1"].Value = "Loan Maturity";
                dataSheet.Cells["C1"].Value = "Amount";
                dataSheet.Cells["A2"].Value = "Fund A";
                dataSheet.Cells["B2"].Value = new DateTime(2024, 3, 17);
                dataSheet.Cells["C2"].Value = 4000000;
                dataSheet.Cells["A3"].Value = "Fund B";
                dataSheet.Cells["B3"].Value = new DateTime(2024, 12, 19);
                dataSheet.Cells["C3"].Value = 2000000;

                // Define the range of the data
                var dataRange = dataSheet.Cells["A1:C3"];

                // Add a new worksheet for the pivot table
                ExcelWorksheet pivotSheet = package.Workbook.Worksheets.Add("PivotTable");

                // Create the pivot table, with the top-left corner at A1
                var pivotTable = pivotSheet.PivotTables.Add(pivotSheet.Cells["A10"], dataRange, "PivotTable");

                // Set row and column fields
                pivotTable.RowFields.Add(pivotTable.Fields["Loan Maturity"]); // Add row field
                pivotTable.DataFields.Add(pivotTable.Fields["Amount"]); // Add data field
                pivotTable.PageFields.Add(pivotTable.Fields["Fund"]);

                //Select the default value for filter
                var defaultfilterValue = "Fund A";
                var pageField = pivotTable.Fields["Fund"];
                pivotTable.PageFields.Add(pageField);

                pageField.Items.Refresh();

                var index = 0;
                foreach (var item in pageField.Items)
                {
                    if (defaultfilterValue.Equals(item.Value.ToString(), StringComparison.CurrentCultureIgnoreCase))
                        break;
                    index++;
                }
                pageField.Items.SelectSingleItem(index);

                // Save to a file (or stream for Azure Blob Storage)
                using (var fileStream = new FileStream(@"EpPlusDefaultSelect.xlsx", FileMode.Create))
                {
                    package.SaveAs(fileStream);
                }
            }
        }
    }
}
