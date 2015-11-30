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
using System.Windows.Navigation;
using System.Windows.Shapes;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace HelloWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string Filename;
        public string SupplierName;

        public MainWindow()
        {
            InitializeComponent();
        }

        //private void PnlMainGrid_OnMouseUpnlMainGrid_MouseUp(object sender, MouseButtonEventArgs e)
        //{
        //    MessageBox.Show("You Clicked me at " + e.GetPosition(this).ToString());
        //}

        private void SupplierNameBox_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var name = SupplierNameBox.Text;
            SupplierName = name;
        }

        private void browsebtn_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog();

            dlg.DefaultExt = ".xmls";
            dlg.Filter = "Excel Documents (.xlsx) | *.xlsx";

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                Filename = dlg.FileName;
                FileNameTextBox.Text = Filename;

                //Paragraph paragraph = new Paragraph();
                //paragraph.Inlines.Add(System.IO.File.ReadAllText(filename));
                //FlowDocument document = new FlowDocument(paragraph);
                //FlowDocReader.Document = document;
            }
        }

        private void Processbtn_OnClick(object sender, RoutedEventArgs e)
        {

            var createRentalUnits = new AddRentalUnits();
            createRentalUnits.CreateRental(Filename, SupplierName);

        }

        
    }
}
