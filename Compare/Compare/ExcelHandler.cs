using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Compare.Models;
using System.Collections;

namespace Compare
{
    public class ExcelHandler
    {
        public ExcelHandler(string filePath, string lookFolderFilePath)
        {
            Globals.ExcelFilePath = filePath;
            Globals.FileNames = GetFilesNames(lookFolderFilePath);
            Globals.ImageLookupFolder = lookFolderFilePath;
        }

        private List<string> GetFilesNames(string lookFolderFilePath)
        {
            var result = new List<string>();
            var fileNames = Directory.GetFiles(lookFolderFilePath).ToList();
            foreach (var fileName in fileNames)
            {
                var name = Path.GetFileNameWithoutExtension(fileName);
                result.Add(name);
            }

            return result;
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
                            person.Path = Path.Combine(Globals.ImageLookupFolder, person.FullName) + ".jpg";

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

            try
            {
                using (var file = File.Open(Globals.ExcelFilePath, FileMode.Open))
                {
                    using (var package = new ExcelPackage(file))
                    {
                        var workbook = package.Workbook;
                        var worksheet = workbook.Worksheets[ConfigurationManager.AppSettings["OrderListWorkSheet"]];
                        var end = worksheet.Dimension.End;

                        for (int i = 1; i < end.Row; i++)
                        {
                            var failed = true;
                            int orderNumberIsNumber = 0;
                            if (!int.TryParse(worksheet.Cells[i + 1, 1].Value.ToString().Trim(), out orderNumberIsNumber))
                                continue;

                            if (worksheet.Cells[i + 1, 1].Value == null)
                                break;

                            var order = new Order
                            {
                                Firstname = worksheet.Cells[i + 1, 6].Value?.ToString().Trim(),
                                Lastname = worksheet.Cells[i + 1, 7].Value?.ToString().Trim(),
                                ChosenImage = worksheet.Cells[i + 1, 11].Value == null ? 0 : Int32.Parse(worksheet.Cells[i + 1, 11].Value.ToString().Trim()),
                                Email = worksheet.Cells[i + 1, 8].Value?.ToString().Trim(),
                                TotalAmount = double.Parse(worksheet.Cells[i + 1, 5].Value.ToString().Trim()),
                                OrderProduct = worksheet.Cells[i + 1, 10].Value?.ToString().Trim(),
                                OrderColor = worksheet.Cells[i + 1, 13].Value?.ToString().Trim(),
                                OrderImagePackage = worksheet.Cells[i + 1, 12].Value?.ToString().Trim(),
                                OrderExtraImagePackage = worksheet.Cells[i + 1, 14].Value?.ToString().Trim()
                            };

                            if (order.TotalAmount < Int32.Parse(ConfigurationManager.AppSettings["MinimumOrderAmount"]))
                            {
                                order.Status = Enums.MailSentStatus.BelowAmount;
                            }
                            else if (order.ChosenImage == 0)
                            {
                                order.Status = Enums.MailSentStatus.ChosenImageNumberMissing;
                            }
                            else
                                failed = false;

                            if (!failed)
                            {
                                var imageFound = false;
                                var image = FileHandler.GetImageFileForOrder(order, Globals.ImageLookupFolder);
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
                }
            }
            catch (Exception e)
            {
                throw;
            }
        }

        internal static bool SaveItemsToExcelFile(List<Person> persons, string file)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(file)))
                {
                    var wb = package.Workbook;
                    var ws = wb.Worksheets.Add("Sheet1");

                    for (var i = 0; i < persons.Count; i++)
                    {
                        var person = persons[i];
                        ws.Cells[i + 1, 1].Value = person.Name;
                        ws.Cells[i + 1, 2].Value = person.Surname;
                        ws.Cells[i + 1, 3].Value = person.Town;
                        ws.Cells[i + 1, 4].Value = person.Path;
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
    }
}
