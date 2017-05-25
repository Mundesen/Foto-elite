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
                return;

            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx";
            var dialogResult = openFileDialog.ShowDialog();
            
            if (dialogResult == System.Windows.Forms.DialogResult.OK)
            {
                var filePath = openFileDialog.FileName;
                var lookupFolder = LookupFolderLabel.Content as string;
                var excelHandler = new ExcelHandler(filePath, lookupFolder);

                var items = excelHandler.StartExcelWork();
                foreach(var item in items)
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
            System.Windows.MessageBox.Show("BØH!");
        }
    }
}
