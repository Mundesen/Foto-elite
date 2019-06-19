using Compare.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using Compare.Helpers;

namespace Compare
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public List<string> ExcelItems = new List<string>();
        public MainWindow()
        {
            InitializeComponent();

            if (Globals.Production)
            {
                System.Windows.MessageBox.Show("Programmet kører i Production mode!","Production mode!",MessageBoxButton.OK,MessageBoxImage.Warning);
            }
        }

        private void ReadExcelPhotoFile_Click(object sender, RoutedEventArgs e)
        {
            RightMainWindow.ItemsSource = null;

            if (string.IsNullOrEmpty(ImageLookupFolderLabel.Content as string))
            {
                System.Windows.MessageBox.Show("Choose folder");
                return;
            }

            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            var dialogResult = openFileDialog.ShowDialog();
            
            if (dialogResult == System.Windows.Forms.DialogResult.OK)
            {
                var filePath = openFileDialog.FileName;
                var lookupFolder = ImageLookupFolderLabel.Content as string;
                var excelHandler = new ExcelHelper(filePath, lookupFolder);

                var allItems = excelHandler.StartReadExcelPhotoFile(false);

                LeftMainWindow.ItemsSource = allItems;
                LeftMainWindow.DisplayMemberPath = "FullName";
                
                var missingItems = excelHandler.StartReadExcelPhotoFile(true);
                RightMainWindow.ItemsSource = missingItems;
                RightMainWindow.DisplayMemberPath = "FullName";
            }
        }

        private void ReadOrderList_Click(object sender, RoutedEventArgs e)
        {
            RightMainWindow.ItemsSource = null;

            if (string.IsNullOrEmpty(ImageLookupFolderLabel.Content as string))
            {
                System.Windows.MessageBox.Show("Choose folder");
                return;
            }

            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            var dialogResult = openFileDialog.ShowDialog();

            if (dialogResult == System.Windows.Forms.DialogResult.OK)
            {
                var filePath = openFileDialog.FileName;
                var imageLookupFolder = ImageLookupFolderLabel.Content as string;
                var excelHandler = new ExcelHelper(filePath, imageLookupFolder);

                try
                {
                    excelHandler.StartReadExcelOrderListFile();

                    //Vis resultat
                    LeftMainWindow.ItemsSource = Globals.ExcelOrderListReady.Select(x => x.Firstname + " " + x.Lastname + " [" + x.Status + "]");
                    RightMainWindow.ItemsSource = Globals.ExcelOrderListFailed.Select(x => x.Firstname + " " + x.Lastname + " [" + x.Status + "]");
                }
                catch(Exception ex)
                {
                    System.Windows.MessageBox.Show(string.Format("Der skete følgende fejl: {0}", ex.Message));
                }
            }
        }

        private void SendMails_Click(object sender, RoutedEventArgs e)
        {
            var result = System.Windows.MessageBox.Show("Are you sure you want to send emails?", "Send Emails", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    if (Globals.ExcelOrderListReady.Count == 0)
                    {
                        System.Windows.MessageBox.Show("No orders ready!");
                        return;
                    }
                    new Views.MailBodySelectorWindow().ShowDialog();
                    if (File.ReadAllText(Globals.SelectedMailBody).Contains("{schoolUrl}"))
                        new Views.SchoolUrlView().ShowDialog();

                    var mailsSend = MailHandler.SendMails();

                    //Vis resultat
                    LeftMainWindow.ItemsSource = Globals.ExcelOrderListReady.Select(x => x.Firstname + " " + x.Lastname + " [" + x.Status + "]");
                    RightMainWindow.ItemsSource = Globals.ExcelOrderListFailed.Select(x => x.Firstname + " " + x.Lastname + " [" + x.Status + "]");

                    System.Windows.MessageBox.Show(string.Format("{0}/{1} mails send!", mailsSend, Globals.ExcelOrderListReady.Count));
                }
                catch(Exception ex)
                {
                    System.Windows.MessageBox.Show(string.Format("Der skete følgende fejl: {0}", ex.Message));
                }
            }
        }

        private void LookupFolderClick(object sender, RoutedEventArgs e)
        {
            var fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ImageLookupFolderLabel.Content = fbd.SelectedPath;
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog();
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            dialog.OverwritePrompt = true;
            var result = dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                var file = dialog.FileName;

                if (File.Exists(file))
                    File.Delete(file);
                
                var persons = GetListBoxItems(RightMainWindow.Items);

                bool saved = ExcelHelper.SaveItemsToExcelFile(persons, file);

                if (saved)
                    System.Windows.Forms.MessageBox.Show("File saved!");
                else
                    System.Windows.Forms.MessageBox.Show("File NOT saved!");
            }            
        }

        private List<Person> GetListBoxItems(ItemCollection items)
        {
            var result = new List<Person>();

            foreach (var item in items)
            {
                result.Add((item as Person));
            }

            return result;
        }

        private void Button_Click1(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(ImageLookupFolderLabel.Content as string))
            {
                System.Windows.MessageBox.Show("Choose folder");
                return;
            }

            var result = System.Windows.MessageBox.Show("Are you sure you want to rename all files in this folder?","Rename", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                var failedFiles = FileHelper.RenameFiles(ImageLookupFolderLabel.Content.ToString());
                if (failedFiles.Count == 0)
                    System.Windows.MessageBox.Show("Files renamed!");
                else
                {
                    var resultString = "";
                    foreach(var file in failedFiles)
                    {
                        resultString += file + Environment.NewLine;
                    }
                    System.Windows.MessageBox.Show("The following files where not renamed: " + Environment.NewLine + Environment.NewLine + resultString);
                }
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            RightMainWindow.ItemsSource = null;
            LeftMainWindow.ItemsSource = null;
            ImageLookupFolderLabel.Content = null; 
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog();
            dialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            dialog.OverwritePrompt = true;
            var result = dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                var file = dialog.FileName;

                if (File.Exists(file))
                    File.Delete(file);

                bool saved = ExcelHelper.SaveItemsToExcelFile(GetListBoxItems(LeftMainWindow.Items), file);

                if (saved)
                    System.Windows.Forms.MessageBox.Show("File saved!");
                else
                    System.Windows.Forms.MessageBox.Show("File NOT saved!");
            }
        }

        private void SortFiles_Click(object sender, RoutedEventArgs e)
        {
            //Diverse tjeks
            if (string.IsNullOrEmpty(Globals.ImageLookupFolder) || Globals.ExcelOrderListReady.Count == 0)
            {
                System.Windows.MessageBox.Show("Billede-mappe ikke valgt, eller ordrer ikke indlæst!");
                return;
            }

            FileHelper.SortFilesAndCreateFolders();

            System.Windows.MessageBox.Show("Filerne og mapper er sorteret.");
        }
    }

}
