using Compare.Models;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;

namespace Compare.Helpers
{
    public class ExcelHelper
    {
        public ExcelHelper(string filePath, string lookFolderFilePath)
        {
            Globals.ExcelFilePath = filePath;
            Globals.FileNames = FileHelper.GetFileNames(lookFolderFilePath);
            Globals.ImageLookupFolder = lookFolderFilePath;
        }

        internal List<Person> StartReadExcelPhotoFile(bool onlyMissingPersons)
        {
            var result = new List<Person>();

            try
            {
                using (var file = File.Open(Globals.ExcelFilePath, FileMode.Open))
                {
                    using (var package = new ExcelPackage(file))
                    {
                        var workbook = package.Workbook;
                        var worksheet = workbook.Worksheets[ConfigurationManager.AppSettings["PhotoFileWorkSheet"]];
                        var end = worksheet.Dimension.End;

                        for (int i = 1; i < end.Row; i++)
                        {
                            var person = new Person
                            {
                                Name = worksheet.Cells[i + 1, 3].Value?.ToString().Trim(),
                                Surname = worksheet.Cells[i + 1, 4].Value?.ToString().Trim(),
                                Town = worksheet.Cells[i + 1, 5].Value?.ToString().Trim()
                            };
                            person.FullName = person.Name + " " + person.Surname;
                            person.ImagePath = Path.Combine(Globals.ImageLookupFolder, person.FullName) + ".jpg";

                            if (onlyMissingPersons)
                            {
                                if (!Globals.FileNames.Exists(p => p.Replace(" ", "").ToLower().Contains(person.Name?.ToLower().Replace(" ", "") + person.Surname?.ToLower().Replace(" ", ""))))
                                {
                                    result.Add(person);
                                }
                            }
                            else
                            {
                                result.Add(person);
                            }
                        }
                    }
                }
                return result;
            }
            catch (Exception e)
            {
                throw;
            }
        }

        internal void StartReadExcelOrderListFile()
        {
            Globals.ExcelOrderListReady = new List<Order>();
            Globals.ExcelOrderListFailed = new List<Order>();

            var data = GetExcelData();

            foreach (var item in data)
            {
                var failed = true;
                var order = new Order
                {
                    FullName = item[ExcelHeaders.Headers.Elevnavn],
                    ChosenImage = item[ExcelHeaders.Headers.Elevnavn],
                    Email = item[ExcelHeaders.Headers.Email],
                    TotalAmount = Convert.ToDecimal(item[ExcelHeaders.Headers.Total].Replace("kr.", "")),
                    OrderProduct = item[ExcelHeaders.Headers.HvadVilDuBestille4],
                    OrderImagePackage = item[ExcelHeaders.Headers.Billedepakke285kr] ?? item[ExcelHeaders.Headers.Billedepakke185kr] ?? item[ExcelHeaders.Headers.AarsbilledeNavn],
                    OrderImagePackageColor = item[ExcelHeaders.Headers.Billedepakke285krFarveEllerSortHvid] ?? item[ExcelHeaders.Headers.Billedepakke185krFarveEllerSortHvid],
                    OrderExtraImagePackage = item[ExcelHeaders.Headers.EkstraBilledepakke],
                    OrderExtraImagePackageColor = item[ExcelHeaders.Headers.EkstraBilledepakkeFarveEllerSortHvid]
                };

                if (order.TotalAmount < Int32.Parse(ConfigurationManager.AppSettings["MinimumOrderAmount"]))
                {
                    order.Status = Enums.MailSentStatus.BelowAmount;
                }
                else if (order.ChosenImage == null)
                {
                    order.Status = Enums.MailSentStatus.ChosenImageNumberMissing;
                }
                else
                    failed = false;

                if (!failed)
                {
                    var imageFound = false;
                    var image = FileHelper.GetImageFileForOrder(order, Globals.ImageLookupFolder);
                    if (!string.IsNullOrEmpty(image))
                    {
                        order.ImageFile = image;
                        order.Status = Enums.MailSentStatus.ReadyForMail;
                        Globals.ExcelOrderListReady.Add(order);
                    }
                    else
                    {
                        order.Status = Enums.MailSentStatus.ImageNotFound;
                        Globals.ExcelOrderListFailed.Add(order);
                    }
                }
                else
                    Globals.ExcelOrderListFailed.Add(order);
            }
        }

        internal static bool SaveItemsToExcelFile(List<Person> persons, string file)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(file)))
                {
                    var wb = package.Workbook;
                    var ws = wb.Worksheets.Add(ConfigurationManager.AppSettings["PhotoFileWorkSheet"]);

                    ws.Cells[1, 1].Value = "Fornavn";
                    ws.Cells[1, 2].Value = "efternavn";
                    ws.Cells[1, 3].Value = "by";
                    ws.Cells[1, 4].Value = "@billede";

                    for (var i = 1; i < persons.Count + 1; i++)
                    {
                        var row = i + 1;
                        var person = persons[i - 1];
                        ws.Cells[row, 1].Value = person.Name;
                        ws.Cells[row, 2].Value = person.Surname;
                        ws.Cells[row, 3].Value = person.Town;
                        ws.Cells[row, 4].Value = person.ImagePath;
                    }

                    package.Save();

                    return true;
                }
            }
            catch (Exception e)
            {
                return false;
            }
        }

        private List<Dictionary<ExcelHeaders.Headers, string>> GetExcelData()
        {
            var prettyfiedData = new List<Dictionary<ExcelHeaders.Headers, string>>();
            using (var file = File.Open(Globals.ExcelFilePath, FileMode.Open))
            {
                using (var package = new ExcelPackage(file))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets[1];
                    var end = worksheet.Dimension.End;

                    for (int i = 2; i <= end.Row; i++)
                    {
                        if(worksheet.Cells[i, 1].Value == null)
                        {
                            break;
                        }

                        var prettyfiedRow = new Dictionary<ExcelHeaders.Headers, string>();
                        for (int j = 0; j < ExcelHeaders.AllHeaders.Length; j++)
                        {
                            var value = worksheet.Cells[i, j + 1].Value?.ToString();
                            if (value != null)
                            {
                                value = ReplaceSpecialCharacters(value.Trim());
                            }
                            prettyfiedRow.Add(ExcelHeaders.AllHeaders[j], value);
                        }
                        prettyfiedData.Add(prettyfiedRow);
                    }
                }
            }

            return prettyfiedData;
        }

        private string ReplaceSpecialCharacters(string input)
        {
            var replacements = new Dictionary<string, string>();
            replacements.Add("\"", "");
            replacements.Add("C8", "ø");
            replacements.Add("C& ", "æ");
            replacements.Add("|", "-");
            replacements.Add("/", "-");

            foreach (var r in replacements)
            {
                input = input.Replace(r.Key, r.Value);
            }

            return input;
        }

        private List<Dictionary<ExcelHeaders.Headers, string>> RemoveDuplicates(List<Dictionary<ExcelHeaders.Headers, string>> data)
        {
            foreach (var item in data)
            {
                if (!string.IsNullOrEmpty(item[ExcelHeaders.Headers.Email]))
                {
                    if (data.Count(x => x[ExcelHeaders.Headers.Email] == item[ExcelHeaders.Headers.Email]) > 1)
                    {

                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(item[ExcelHeaders.Headers.MobilForaeldre]))
                    {
                        if (data.Count(x => x[ExcelHeaders.Headers.MobilForaeldre] == item[ExcelHeaders.Headers.MobilForaeldre]) > 1)
                        {

                        }
                    }
                }
            }

            return data;
        }

    }
}
