using Application.Exceptions;
using Application.Parser;
using Application.Writers;
using Microsoft.Win32;
using System.IO;
using System.Windows;

namespace Application.UI.UserControls
{
    internal class FormFillingManager
    {
        public void FullOnePieceFile(int firstLine, String formPath, int designLine, Parser.Parser parser)
        {
            String fileToParse = "";
            String fileToSave = "";

            try
            {
                fileToParse = this.getFileToOpen();
                if (fileToParse == "") return;

                List<Data.Piece> data = parser.ParseFile(fileToParse);
                Dictionary<string, string> header = parser.GetHeader();

                fileToSave = this.getFileToSave();
                if (fileToSave == "") return;

                OnePieceWriter excelWriter = new OnePieceWriter(fileToSave, firstLine, formPath);
                excelWriter.WriteHeader(header, designLine);
                excelWriter.WriteData(data);
            }
            catch (MeasureTypeNotFoundException e)
            {
                this.displayError(e.Message);
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

        public void FullFivePieesFile(Parser.Parser parser)
        {
            List<Data.Piece>? data;
            if (parser is TextFileParser)
            {
                data = this.getDataFromFolder(parser);
            }
            else
            {
                data = parser.ParseFile(this.getFileToOpen());
            }

            if (data == null) return;

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

        private List<Data.Piece>? getDataFromFolder(Parser.Parser parser)
        {
            List<Data.Piece> data = new List<Data.Piece>();

            String folderName = this.getFolderToOpen();
            if (folderName == "") return null;

            DirectoryInfo directory = new DirectoryInfo(folderName);

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
                }
                catch (MeasureTypeNotFoundException)
                {
                    this.displayError("Un type de mesure n'a pas été trouvé dans le fichier " + file.FullName);
                }
            }

            return data;
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
