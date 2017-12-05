﻿using Compare.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Compare
{
    class FileHandler
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

            return images.FirstOrDefault(x => Path.GetFileNameWithoutExtension(x).Split(' ')[0].Replace("_", "").Equals(order.ChosenImage.ToString()));
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
                var currentFolder = Path.Combine(rootFolder, currentOrderNumber, currentOrder.OrderColor.Replace("/", ""), string.IsNullOrEmpty(currentOrder.OrderImagePackage) ? currentOrder.OrderProduct : currentOrder.OrderImagePackage);
                
                Directory.CreateDirectory(currentFolder);

                File.Copy(currentOrder.ImageFile, Path.Combine(currentFolder, Path.GetFileName(currentOrder.ImageFile)));

                if (!string.IsNullOrEmpty(currentOrder.OrderExtraImagePackage))
                {
                    var extraOrderFolder = Path.Combine(rootFolder, currentOrderNumber, currentOrder.OrderExtraImagePackage);
                    Directory.CreateDirectory(extraOrderFolder);
                    File.Copy(currentOrder.ImageFile, Path.Combine(extraOrderFolder, Path.GetFileName(currentOrder.ImageFile)));
                }
            }
        }
    }
}