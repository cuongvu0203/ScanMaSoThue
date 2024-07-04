using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ScanMaSoThue
{
    public partial class Form1 : Form
    {
        private static readonly HttpClient client = new HttpClient();

        public Form1()
        {
            InitializeComponent();
            progressBar.Minimum = 0;
            progressBar.Value = 0;

            // Set the license context for EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private async void btnSelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls;*.xlsx",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                if (!filePath.EndsWith(".xls") && !filePath.EndsWith(".xlsx"))
                {
                    MessageBox.Show("File không đúng định dạng", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                List<string> idNumbers = ExtractIdNumbersFromExcel(filePath);

                if (idNumbers == null || idNumbers.Count == 0)
                {
                    MessageBox.Show("Không tìm thấy số căn cước trong file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                progressBar.Maximum = idNumbers.Count;
                progressBar.Value = 0;

                List<Dictionary<string, string>> results = new List<Dictionary<string, string>>();

                foreach (var id in idNumbers)
                {
                    var result = await SearchInfo(id);
                    results.Add(result);
                    progressBar.Value += 1;
                }

                dataGridView.DataSource = results.Select(d => new
                {
                    CompanyName = d.GetValueOrDefault("companyName"),
                    TaxCode = d.GetValueOrDefault("taxCode"),
                    Address = d.GetValueOrDefault("address"),
                    Owner = d.GetValueOrDefault("owner"),
                    OperatingDay = d.GetValueOrDefault("operatingDay"),
                    ManagedBy = d.GetValueOrDefault("managedBy"),
                    Status = d.GetValueOrDefault("status"),
                    DateUpdate = d.GetValueOrDefault("dateUpdate")
                }).ToList();

                SaveResultsToExcel(results, "output.xlsx");
                MessageBox.Show("Processing completed. Results saved to output.xlsx.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private List<string> ExtractIdNumbersFromExcel(string filePath)
        {
            List<string> idNumbers = new List<string>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 1; row <= rowCount; row++)
                {
                    var cellValue = worksheet.Cells[row, 1].Text;
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        idNumbers.Add(cellValue.Trim());
                    }
                }
            }

            return idNumbers;
        }

        private async Task<Dictionary<string, string>> SearchInfo(string req)
        {
            var extractedData = new Dictionary<string, string>();

            try
            {
                var request = new HttpRequestMessage(HttpMethod.Get, $"https://masothue.com/Search/?q={req}&type=auto");
                request.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36");

                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();  // Throws an exception if the HTTP response status is an error code.

                var responseBody = await response.Content.ReadAsStringAsync();

                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(responseBody);

                // Extract company name
                var companyName = htmlDoc.DocumentNode.SelectSingleNode("//th/span[contains(@class, 'copy')]")?.InnerText.Trim();
                extractedData["companyName"] = companyName;

                // Extract table rows
                var rows = htmlDoc.DocumentNode.SelectNodes("//table[@class='table-taxinfo']//tr");
                if (rows != null)
                {
                    foreach (var row in rows)
                    {
                        var cells = row.SelectNodes("td");
                        if (cells != null && cells.Count == 2)
                        {
                            var key = cells[0].InnerText.Trim();
                            var valueNode = cells[1].SelectSingleNode(".//span[contains(@class, 'copy')]") ?? cells[1];
                            var value = valueNode.InnerText.Trim();

                            switch (key)
                            {
                                case "Mã số thuế cá nhân":
                                    extractedData["taxCode"] = value;
                                    break;
                                case "Địa chỉ":
                                    extractedData["address"] = value;
                                    break;
                                case "Người đại diện":
                                    extractedData["owner"] = value.Replace(" Ẩn thông tin", "");
                                    break;
                                case "Ngày hoạt động":
                                    extractedData["operatingDay"] = value;
                                    break;
                                case "Quản lý bởi":
                                    extractedData["managedBy"] = value;
                                    break;
                                case "Tình trạng":
                                    extractedData["status"] = value;
                                    break;
                            }
                        }
                    }
                }

                // Extract update time
                var updateTime = htmlDoc.DocumentNode.SelectSingleNode("//td/em")?.InnerText.Trim();
                extractedData["dateUpdate"] = updateTime;
            }
            catch (HttpRequestException ex) when (ex.StatusCode == System.Net.HttpStatusCode.Forbidden)
            {
                Console.WriteLine("Access forbidden: Check if you are being blocked by the server.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Đã xảy ra lỗi:", ex);
            }

            return extractedData;
        }

        private void SaveResultsToExcel(List<Dictionary<string, string>> results, string filePath)
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Results");

                worksheet.Cells[1, 1].Value = "CompanyName";
                worksheet.Cells[1, 2].Value = "TaxCode";
                worksheet.Cells[1, 3].Value = "Address";
                worksheet.Cells[1, 4].Value = "Owner";
                worksheet.Cells[1, 5].Value = "OperatingDay";
                worksheet.Cells[1, 6].Value = "ManagedBy";
                worksheet.Cells[1, 7].Value = "Status";
                worksheet.Cells[1, 8].Value = "DateUpdate";

                for (int i = 0; i < results.Count; i++)
                {
                    var result = results[i];
                    worksheet.Cells[i + 2, 1].Value = result.GetValueOrDefault("companyName");
                    worksheet.Cells[i + 2, 2].Value = result.GetValueOrDefault("taxCode");
                    worksheet.Cells[i + 2, 3].Value = result.GetValueOrDefault("address");
                    worksheet.Cells[i + 2, 4].Value = result.GetValueOrDefault("owner");
                    worksheet.Cells[i + 2, 5].Value = result.GetValueOrDefault("operatingDay");
                    worksheet.Cells[i + 2, 6].Value = result.GetValueOrDefault("managedBy");
                    worksheet.Cells[i + 2, 7].Value = result.GetValueOrDefault("status");
                    worksheet.Cells[i + 2, 8].Value = result.GetValueOrDefault("dateUpdate");
                }

                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);
            }
        }
    }
}
