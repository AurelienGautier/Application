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
        public void FullOnePieceFile(int firstLine, String formPath, int designLine, Parser.Parser parser, bool sign, bool modify)
        {
            try
            {
                String fileToParse = "";
                String fileToSave = "";

                fileToParse = this.GetFileToOpen();
                if (fileToParse == "") return;

                List<Data.Piece> data = parser.ParseFile(fileToParse);
                Dictionary<string, string> header = parser.GetHeader();

                fileToSave = this.GetFileToSave();
                if (fileToSave == "") return;

                OnePieceWriter excelWriter = new OnePieceWriter(fileToSave, firstLine, formPath, modify);
                excelWriter.WriteHeader(header, designLine);
                excelWriter.WriteData(data, sign);
            }
            catch (MeasureTypeNotFoundException e)
            {
                this.displayError(e.Message);
            }
            catch (IncorrectFormatException e)
            {
                this.displayError(e.Message);
            }
            catch (ExcelFileAlreadyInUseException e)
            {
                this.displayError(e.Message);
            }
        }

        public void FullFivePiecesFile(String formToModify, Parser.Parser parser, bool sign, bool modify)
        {
            List<Data.Piece>? data;
            if (parser is TextFileParser)
            {
                data = this.getDataFromFolder(parser);
            }
            else
            {
                data = parser.ParseFile(this.GetFileToOpen());
            }

            if (data == null) return;

            Dictionary<string, string> header = parser.GetHeader();

            String fileToSave = this.GetFileToSave();
            if (fileToSave == "") return;

            try
            {
                FivePiecesWriter excelWriter = new FivePiecesWriter(fileToSave, formToModify, modify);
                excelWriter.WriteHeader(header, 25);
                excelWriter.WriteData(data, sign);
            }
            catch (ExcelFileAlreadyInUseException)
            {
                this.displayError("Le fichier excel est déjà en cours d'utilisation");
            }
            catch (Exception e)
            {
                this.displayError(e.Message);
            }
        }

        private List<Data.Piece>? getDataFromFolder(Parser.Parser parser)
        {
            String folderName = this.getFolderToOpen();
            if (folderName == "") return null;

            DirectoryInfo directory = new DirectoryInfo(folderName);

            List<Data.Piece> data = directory.GetFiles()
                .Select(file => file.FullName)
                .SelectMany(fileName =>
                {
                    try
                    {
                        return parser.ParseFile(fileName);
                    }
                    catch (IncorrectFormatException)
                    {
                        this.displayError("Le format du fichier " + fileName + " est incorrect.");
                        return Enumerable.Empty<Data.Piece>();
                    }
                    catch (MeasureTypeNotFoundException)
                    {
                        this.displayError("Un type de mesure n'a pas été trouvé dans le fichier " + fileName);
                        return Enumerable.Empty<Data.Piece>();
                    }
                })
                .ToList();

            return data;
        }

        public String GetFileToOpen()
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

        public String GetFileToSave()
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
