using Application.Exceptions;
using Application.Writers;
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
                Parser parser = new Parser();
                List<Piece> data = parser.ParseFile(fileToParse);

                OnePieceWriter excelWriter = new OnePieceWriter(fileToSave);
                excelWriter.WriteData(data);
            }
            catch (IncorrectFormatException)
            {
                this.displayError("Le format du fichier est incorrect.");
            }
            catch (ExcelFileAlreadyInUseException)
            {
                this.displayError("Le fichier excel est déjà en cours d'utilisation");
            }

            /*FreeConsole();*/
        }

        private void Rapport5Pieces_Click(object sender, RoutedEventArgs e)
        {
            /*AllocConsole();*/

            String folderName = this.getFolderToOpen();
            if (folderName == "") return;

            DirectoryInfo directory = new DirectoryInfo(folderName);
            if(!directory.Exists) return;

            Parser parser = new Parser();
            List<Piece> data = new List<Piece>();

            // Parsing de tous les fichiers du répertoire
            foreach(FileInfo file in directory.GetFiles())
            {
                try
                {
                    data.AddRange(parser.ParseFile(file.FullName));
                }
                catch (IncorrectFormatException) 
                {
                    this.displayError("Le format du fichier " + file.FullName + " est incorrect.");
                    return;
                }
            }

            String fileToSave = this.getFileToSave();
            if (fileToSave == "") return;

            try
            {
                FivePiecesWriter excelWriter = new FivePiecesWriter(fileToSave);
                excelWriter.WriteData(data);
            }
            catch (ExcelFileAlreadyInUseException)
            {
                this.displayError("Le fichier excel est déjà en cours d'utilisation");
            }

            /*FreeConsole();*/
        }


        private String getFileToOpen()
        {
            var dialog = new OpenFileDialog();

            String fileName = "";

            if (dialog.ShowDialog() == true) fileName = dialog.FileName;

            return fileName;
        }

        private String getFolderToOpen()
        {
            var dialog = new OpenFolderDialog();

            String folderName = "";

            if (dialog.ShowDialog() == true) folderName = dialog.FolderName;

            return folderName;
        }

        private String getFileToSave()
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Fichiers Excel (*.xlsx)|*.xlsx";
            saveFileDialog.FileName = "rapport";

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
            String caption = "Erreur";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Error;
            MessageBoxResult result;

            result = MessageBox.Show(errorMessage, caption, button, icon, MessageBoxResult.Yes);
        }
    }
}