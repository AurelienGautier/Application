using Application.Exceptions;
using System;
using System.IO;
using System.Runtime.InteropServices;
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

namespace Application
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        [DllImport("Kernel32")]
        public static extern void AllocConsole();

        [DllImport("Kernel32", SetLastError = true)]
        public static extern void FreeConsole();

        public MainWindow()
        {
            InitializeComponent();
        }

        public void Rapport1Piece_Click(object sender, RoutedEventArgs e)
        {
            /*AllocConsole();*/

            String fileName = this.getFileToOpen();

            if (fileName == "") return;

            try
            {
                Parser parser = new Parser(fileName);
                List<Piece> data = parser.ParseFile();

                ExcelWriter excelWriter = new ExcelWriter();
                excelWriter.WriteData(data);
            }
            catch(IncorrectFormatException)
            {
                this.displayError();
            }

            /*FreeConsole();*/
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private string getFileToOpen()
        {
            // Configure open file dialog box
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.FileName = "Document"; // Default file name
            dialog.DefaultExt = ".txt"; // Default file extension

            // Show open file dialog box
            bool? result = dialog.ShowDialog();

            string fileName = "";

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                fileName = dialog.FileName;
            }

            return fileName;
        }

        private void displayError()
        {
            string messageBoxText = "Le format du fichier est incorrect.";
            string caption = "Erreur";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Error;
            MessageBoxResult result;

            result = MessageBox.Show(messageBoxText, caption, button, icon, MessageBoxResult.Yes);
        }
    }
}