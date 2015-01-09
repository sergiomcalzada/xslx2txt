using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using xslx2textlib;

namespace xslx2text
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly FileFinder finder = new FileFinder();

        private readonly Xml2TxtParser parser = new Xml2TxtParser();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            var files = finder.FindFiles(AppDomain.CurrentDomain.BaseDirectory, "*.xlsx", true);
            foreach (var file in files)
            {
                parser.Parse(file);
            }
           
        }

        private void btnFindFile_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            var dlg = new Microsoft.Win32.OpenFileDialog()
                          {
                              DefaultExt = ".xslx",
                          };



            // Set filter for file extension and default file extension 

            // Display OpenFileDialog by calling ShowDialog method 
            var result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                txtFileOrFolder.Text = filename;
            }
        }
    }
}
