using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Compare
{
    class ExcelHandler
    {
        private string ExcelFilePath { get; set; }
        private List<string> FileNames { get; set; }

        public ExcelHandler(string filePath, string lookFolderFilePath)
        {
            ExcelFilePath = filePath;
            FileNames = GetFilesNames(lookFolderFilePath);
        }

        private List<string> GetFilesNames(string lookFolderFilePath)
        {
            var result = new List<string>();
            var fileNames = Directory.GetFiles(lookFolderFilePath).ToList();
            foreach(var fileName in fileNames)
            {
                var name = Path.GetFileNameWithoutExtension(fileName);
                result.Add(name);
            }

            return result;
        }

        public List<string> StartExcelWork()
        {
            var result = new List<string>();

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
                            var person = worksheet.Cells[i + 1, 3].Value + " " + worksheet.Cells[i + 1, 4].Value;

                            if (!FileNames.Exists(x => x.Replace(" ", "").ToLower().Contains(person.Replace(" ", "").ToLower())))
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
                return new List<string> { "File is open!" };
            }
        }
    }
}
