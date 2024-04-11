using Application.Exceptions;
using Microsoft.Win32;
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

            String fileToParse = this.getFileToOpen();
            if (fileToParse == "") return;
            String fileToSave = this.getFileToSave();
            if (fileToSave == "") return;

            try
            {
                Parser parser = new Parser(fileToParse);
                List<Piece> data = parser.ParseFile();

                ExcelWriter excelWriter = new ExcelWriter(fileToSave);
                excelWriter.WriteData(data);
            }
            catch(IncorrectFormatException)
            {
                this.displayError("Le format du fichier est incorrect.");
            }
            catch(ExcelFileAlreadyInUse)
            {
                this.displayError("Le fichier excel est déjà en cours d'utilisation");
            }

            /*FreeConsole();*/
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            
        }

        private string getFileToOpen()
        {
            var dialog = new OpenFileDialog();
            dialog.FileName = "Document";
            dialog.DefaultExt = ".txt";

            string fileName = "";

            if (dialog.ShowDialog() == true)
            {
                fileName = dialog.FileName;
            }

            return fileName;
        }

        private String getFileToSave()
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Fichiers Excel (*.xlsx)|*.xlsx";
            saveFileDialog.FileName = "rappport1piece";

            String fileName = "";
         
            if (saveFileDialog.ShowDialog() == true)
            {
                fileName = saveFileDialog.FileName;
            }

            if(fileName.Length > 5)
                fileName = fileName.Remove(fileName.Length - 5);

            return fileName;
        }

        private void displayError(String errorMessage)
        {
            string caption = "Erreur";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Error;
            MessageBoxResult result;

            result = MessageBox.Show(errorMessage, caption, button, icon, MessageBoxResult.Yes);
        }
    }
}