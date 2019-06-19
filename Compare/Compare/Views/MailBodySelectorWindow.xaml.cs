using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Compare.Views
{
    /// <summary>
    /// Interaction logic for MailBodySelectorWindow.xaml
    /// </summary>
    public partial class MailBodySelectorWindow : Window
    {
        public MailBodySelectorWindow()
        {
            InitializeComponent();

            mailBodyComboBox.ItemsSource = ConfigurationManager.AppSettings["PathBodyTemplates"].Split(';');
            mailBodyComboBox.SelectedIndex = 0;
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            Globals.SelectedMailBody = mailBodyComboBox.SelectedItem.ToString();
            this.Close();
        }
    }
}
