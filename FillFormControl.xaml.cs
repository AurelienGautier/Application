using Application.Exceptions;
using Application.Writers;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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

namespace Application
{
    /// <summary>
    /// Logique d'interaction pour FillForm.xaml
    /// </summary>
    public partial class FillFormControl : UserControl
    {
        public FillFormControl()
        {
            InitializeComponent();

            List<string> items = new List<string> {
                "Rapport 1 pièce",
                "Outillage de contrôle",
                "Rapport 5 pièces",
                "Bague lisse",
                "Calibre à machoire",
                "Capabilité",
                "Étalon de colonne de mesure",
                "Tampon lisse"
            };

            Forms.ItemsSource = items;
        }

        private void FillAform(object sender, RoutedEventArgs e)
        {
            switch (Forms.SelectedItem)
            {
                case "Rapport 1 pièce":
                    this.FullOnePieceFile(30, Environment.CurrentDirectory + "\\form\\rapport1piece", 26, 66);
                    break;
                case "Outillage de contrôle":
                    this.FullOnePieceFile(26, Environment.CurrentDirectory + "\\form\\outillageDeControle", 25, 62);
                    break;
                case "Rapport 5 pièces":
                    this.FullFivePieesFile();
                    break;
            }
        }

        public void FullOnePieceFile(int firstLine, String formPath, int designLine, int operatorLine)
        {
            String fileToParse = this.getFileToOpen();
            if (fileToParse == "") return;
            String fileToSave = this.getFileToSave();
            if (fileToSave == "") return;

            try
            {
                Parser parser = new Parser();
                List<Piece> data = parser.ParseFile(fileToParse);
                Dictionary<string, string> header = parser.GetHeader();

                OnePieceWriter excelWriter = new OnePieceWriter(fileToSave, firstLine, formPath);
                excelWriter.WriteHeader(header, designLine);
                excelWriter.WriteData(data);
            }
            catch (MeasureTypeNotFoundException)
            {
                this.displayError("Un type de mesure n'a pas été reconnu dans le fichier " + fileToParse);
            }
            catch (IncorrectFormatException)
            {
                this.displayError("Le format du fichier est incorrect.");
            }
            catch (ExcelFileAlreadyInUseException)
            {
                this.displayError("Le fichier excel est déjà en cours d'utilisation");
            }
        }

        public void FullFivePieesFile()
        {
            String folderName = this.getFolderToOpen();
            if (folderName == "") return;

            DirectoryInfo directory = new DirectoryInfo(folderName);
            if (!directory.Exists) return;

            Parser parser = new Parser();
            List<Piece> data = new List<Piece>();

            // Parsing de tous les fichiers du répertoire
            foreach (FileInfo file in directory.GetFiles())
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
                catch (MeasureTypeNotFoundException)
                {
                    this.displayError("Un type de mesure n'a pas été trouvé dans le fichier " + file.FullName);
                    return;
                }
            }
            Dictionary<string, string> header = parser.GetHeader();

            String fileToSave = this.getFileToSave();
            if (fileToSave == "") return;

            try
            {
                FivePiecesWriter excelWriter = new FivePiecesWriter(fileToSave);
                excelWriter.WriteHeader(header, 25);
                excelWriter.WriteData(data);
            }
            catch (ExcelFileAlreadyInUseException)
            {
                this.displayError("Le fichier excel est déjà en cours d'utilisation");
            }
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

            if (fileName.Length > 5)
                fileName = fileName.Remove(fileName.Length - 5);

            return fileName;
        }

        private void displayError(String errorMessage)
        {
            String caption = "Erreur";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Error;

            MessageBox.Show(errorMessage, caption, button, icon, MessageBoxResult.Yes);
        }
    }
}
