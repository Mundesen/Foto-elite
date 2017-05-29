using System.Windows.Input;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.IO;

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
        }

        private void ReadExcelClick(object sender, RoutedEventArgs e)
        {
            ExcelListBox.Items.Clear();

            if (string.IsNullOrEmpty(LookupFolderLabel.Content as string))
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
                var lookupFolder = LookupFolderLabel.Content as string;
                var excelHandler = new ExcelHandler(filePath, lookupFolder);

                var allItems = excelHandler.StartExcelWork(false);
                foreach (var item in allItems)
                {
                    ListBoxItem itm = new ListBoxItem();
                    itm.Content = item;

                    ExcelListBox1.Items.Add(itm);
                }

                var missingItems = excelHandler.StartExcelWork(true);
                foreach(var item in missingItems)
                {
                    ListBoxItem itm = new ListBoxItem();
                    itm.Content = item;

                    ExcelListBox.Items.Add(itm);
                }
            }
        }

        private void LookupFolderClick(object sender, RoutedEventArgs e)
        {
            var fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                LookupFolderLabel.Content = fbd.SelectedPath;
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
                
                var listBoxItemsAsStrings = GetListBoxItemsAsStrings();

                bool saved = ExcelHandler.SaveItemsToExcelFile(listBoxItemsAsStrings, file);

                if (saved)
                    System.Windows.Forms.MessageBox.Show("File saved!");
                else
                    System.Windows.Forms.MessageBox.Show("File NOT saved!");
            }            
        }

        private List<string> GetListBoxItemsAsStrings()
        {
            var result = new List<string>();

            foreach(var item in ExcelListBox.Items)
            {
                result.Add((item as ListBoxItem).Content.ToString());
            }

            return result;
        }

        private void Button_Click1(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(LookupFolderLabel.Content as string))
            {
                System.Windows.MessageBox.Show("Choose folder");
                return;
            }

            var result = System.Windows.MessageBox.Show("Are you sure you want to rename all files in this folder?","Rename", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                var failedFiles = FileHandler.RenameFiles(LookupFolderLabel.Content.ToString());
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

        private void ExcelListBox1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        
    }

}
