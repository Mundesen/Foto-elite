using System.Windows;

namespace Compare.Views
{
    /// <summary>
    /// Interaction logic for SchoolLinkView.xaml
    /// </summary>
    public partial class SchoolUrlView : Window
    {
        public SchoolUrlView()
        {
            InitializeComponent();

            SchoolUrlTextBox.Text = Globals.SchoolUrl;
        }

        private void SchoolUrlOKButton_Click(object sender, RoutedEventArgs e)
        {
            Globals.SchoolUrl = SchoolUrlTextBox.Text;
            this.Close();
        }

        private void SchoolUrlCancelButton_Click(object sender, RoutedEventArgs e)
        {
            Globals.SchoolUrl = "";
            this.Close();
        }
    }
}
