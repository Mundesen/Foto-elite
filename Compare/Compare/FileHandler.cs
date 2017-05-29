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
            var files = Directory.GetFiles(directory);
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

                newName = CheckForDuplicate(directory, newName, extension);
                try
                {
                    File.Move(file, Path.Combine(directory, newName));
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
            //for (var i = 0; i < Directory.GetFiles(directory).Count(); i++)
            //{
            //    if (File.Exists(file))
            //    {

            //    }
            //}

            return Path.Combine(directory, file + extension);
        }
    }
}