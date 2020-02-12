using Compare.Models;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Compare.Helpers
{
    class FileHelper
    {
        internal static List<string> RenameFiles(string directory)
        {
            var files = Directory.GetFiles(directory, "*", SearchOption.AllDirectories);
            var result = new List<string>();
            foreach (var file in files)
            {
                var extension = Path.GetExtension(file);
                var oldName = Path.GetFileNameWithoutExtension(file);
                var invalidChars = Path.GetInvalidFileNameChars().ToList();
                invalidChars.Add('_');

                var newName = Regex.Replace(oldName, "[0-9]", "");
                foreach (var c in invalidChars)
                {
                    newName = newName.Replace(c.ToString(), string.Empty).Trim();
                }

                newName = CheckForDuplicate(Path.GetDirectoryName(file), newName, extension);
                try
                {
                    File.Move(file, Path.Combine(Path.GetDirectoryName(file), newName));
                }
                catch (IOException)
                {
                    result.Add(oldName + extension);
                }
            }

            return result;
        }

        private static string CheckForDuplicate(string directory, string file, string extension)
        {
            return Path.Combine(directory, file + extension);
        }

        internal static string GetImageFileForOrder(Order order, string imageFilePath)
        {
            var images = Directory.GetFiles(imageFilePath);

            return images.FirstOrDefault(x => Path.GetFileNameWithoutExtension(x).Equals(order.ChosenImage.ToString()));
        }

        internal static void SortFilesAndCreateFolders()
        {
            //Flyt billederne til mapper pr. ordre
            //  - ..[bestillinsliste_path]\SortFolder\OrdreX\[color_folders]
            var rootFolder = Path.Combine(Path.GetDirectoryName(Globals.ExcelFilePath), "SortFolder");
            Directory.CreateDirectory(rootFolder);
            var orderFolderCounter = Directory.GetDirectories(rootFolder).Count() + 1;

            for(int i = 0; i < Globals.ExcelOrderListReady.Count; i++)
            {
                var currentOrder = Globals.ExcelOrderListReady[i];
                var currentOrderNumber = "Order" + orderFolderCounter;
                var currentFolder = Path.Combine(rootFolder, currentOrderNumber, currentOrder.OrderImagePackageColor.Replace("/", ""), string.IsNullOrEmpty(currentOrder.OrderImagePackage) ? currentOrder.OrderProduct : currentOrder.OrderImagePackage);
                
                Directory.CreateDirectory(currentFolder);
                
                var imageFileName = GetNewImageName(currentFolder, currentOrder.ImageFile);
                var destFilePath = Path.Combine(currentFolder, imageFileName);
                
                File.Copy(currentOrder.ImageFile, destFilePath);

                if (!string.IsNullOrEmpty(currentOrder.OrderExtraImagePackage))
                {
                    var extraOrderFolder = Path.Combine(rootFolder, currentOrderNumber, currentOrder.OrderExtraImagePackage.Replace("/",""), currentOrder.OrderExtraImagePackageColor.Replace("/", ""));

                    Directory.CreateDirectory(extraOrderFolder);

                    imageFileName = GetNewImageName(extraOrderFolder, currentOrder.ImageFile);
                    
                    File.Copy(currentOrder.ImageFile, Path.Combine(extraOrderFolder, imageFileName));
                }
            }
        }
        public static List<string> GetFileNames(string lookFolderFilePath)
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

        private static string GetNewImageName(string folder, string imageName)
        {
            return GetNumberOfDuplicateImages(folder, imageName) == 0 ? Path.GetFileName(imageName) : Path.GetFileNameWithoutExtension(Path.GetFileName(imageName)) + "_" + GetNumberOfDuplicateImages(folder, imageName) + Path.GetExtension(imageName);
        }

        private static int GetNumberOfDuplicateImages(string folder, string imageName)
        {
            return Directory.GetFiles(folder).Count(x => Path.GetFileNameWithoutExtension(x).StartsWith(Path.GetFileNameWithoutExtension(imageName)));
        }
    }
}