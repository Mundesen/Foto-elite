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
    class ExcelHandler
    {
        private string ExcelFilePath { get; set; }
        private List<string> FileNames { get; set; }
        private string LookupFolder { get; set; }

        public ExcelHandler(string filePath, string lookFolderFilePath)
        {
            ExcelFilePath = filePath;
            FileNames = GetFilesNames(lookFolderFilePath);
            LookupFolder = lookFolderFilePath;
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

        public List<Person> StartExcelWork(bool onlyMissingPersons)
        {
            var result = new List<Person>();

            try
            {
                using (var file = File.Open(ExcelFilePath, FileMode.Open))
                {
                    using (var package = new ExcelPackage(file))
                    {
                        var workbook = package.Workbook;
                        var worksheet = workbook.Worksheets[ConfigurationManager.AppSettings["WorkSheet"]];
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
                            person.Path = Path.Combine(LookupFolder, person.FullName) + ".jpg";

                            if (onlyMissingPersons)
                            {
                                if (!FileNames.Exists(p => p.Replace(" ", "").ToLower().Contains(person.Name?.ToLower().Replace(" ", "") + person.Surname?.ToLower().Replace(" ", ""))))
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
