using System.Configuration;
using System.Windows;

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
