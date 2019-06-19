using Compare.Models;
using System;
using System.Collections.Generic;
using System.Configuration;

namespace Compare
{
    public class Globals
    {
        public static string ExcelFilePath { get; set; }
        public static List<string> FileNames { get; set; }
        public static string ImageLookupFolder { get; set; }
        public static List<Order> ExcelOrderListReady { get; set; }
        public static List<Order> ExcelOrderListFailed { get; set; }
        public static string SchoolUrl { get; set; }
        public static string SelectedMailBody { get; set; }

        public static bool Production = Boolean.Parse(ConfigurationManager.AppSettings["Production"]);
    }
}
