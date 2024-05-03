using Application.Exceptions;
using Application.Parser;
using Application.Writers;
using Microsoft.Win32;
using System.IO;

namespace Application.UI.UserControls
{
    internal class FormFillingManager
    {
        public void FillOnePieceFile(Data.Form form, Parser.Parser parser)
        {
            try
            {
                String fileToParse = "";
                String fileToSave = "";

                fileToParse = this.GetFileToOpen("Choisir le fichier à convertir", parser.GetFileExtension());
                if (fileToParse == "") return;

                List<Data.Piece> data = parser.ParseFile(fileToParse);
                Dictionary<string, string> header = parser.GetHeader();

                fileToSave = this.GetFileToSave();
                if (fileToSave == "") return;

                OnePieceWriter excelWriter = new OnePieceWriter(fileToSave, form.FirstLine, form.Path, form.Modify);
                excelWriter.WriteHeader(header, form.DesignLine);
                excelWriter.WriteData(data, form.Sign);
            }
            catch (MeasureTypeNotFoundException e)
            {
                MainWindow.DisplayError(e.Message);
            }
            catch (IncorrectFormatException e)
            {
                MainWindow.DisplayError(e.Message);
            }
            catch (ExcelFileAlreadyInUseException e)
            {
                MainWindow.DisplayError(e.Message);
            }
        }

        /*-------------------------------------------------------------------------*/

        public void FillFivePiecesFile(String formToModify, Parser.Parser parser, bool sign, bool modify)
        {
            List<Data.Piece>? data;
            if (parser is TextFileParser)
            {
                data = this.getDataFromFolder(parser);
            }
            else
            {
                String fileToParse = this.GetFileToOpen("Sélectionner le fichier à convertir", "(*.xlsx;*.xlsm)|*.xlsx;*.xlsm");
                if(fileToParse == "") return;

                data = parser.ParseFile(fileToParse);
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
                MainWindow.DisplayError("Le fichier excel est déjà en cours d'utilisation");
            }
        }

        /*-------------------------------------------------------------------------*/

        private List<Data.Piece>? getDataFromFolder(Parser.Parser parser)
        {
            String folderName = this.getFolderToOpen();
            if (folderName == "") return null;

            DirectoryInfo directory = new DirectoryInfo(folderName);

            List<Data.Piece> data = directory
                .GetFiles()
                .Where(file => file.Extension == ".mit" || file.Extension == ".txt" || file.Extension == ".MIT")
                .Select(file => file.FullName)
                .SelectMany(fileName =>
                {
                    try
                    {
                        return parser.ParseFile(fileName);
                    }
                    catch (IncorrectFormatException)
                    {
                        MainWindow.DisplayError("Le format du fichier " + fileName + " est incorrect.");
                        return Enumerable.Empty<Data.Piece>();
                    }
                    catch (MeasureTypeNotFoundException)
                    {
                        MainWindow.DisplayError("Un type de mesure n'a pas été trouvé dans le fichier " + fileName);
                        return Enumerable.Empty<Data.Piece>();
                    }
                })
                .ToList();

            return data;
        }

        /*-------------------------------------------------------------------------*/

        public String GetFileToOpen(String title, String extensions)
        {
            var dialog = new OpenFileDialog();
            dialog.Title = title;
            dialog.Filter = extensions;

            String fileName = "";

            if (dialog.ShowDialog() == true) fileName = dialog.FileName;

            return fileName;
        }

        /*-------------------------------------------------------------------------*/

        private String getFolderToOpen()
        {
            var dialog = new OpenFolderDialog();

            String folderName = "";

            if (dialog.ShowDialog() == true) folderName = dialog.FolderName;

            return folderName;
        }

        /*-------------------------------------------------------------------------*/

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
    }
}
