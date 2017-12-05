using System;
using System.Collections.Generic;
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

namespace Compare
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
