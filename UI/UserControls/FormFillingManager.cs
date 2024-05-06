using Application.Exceptions;
using Application.Parser;
using Application.Writers;
using Microsoft.Win32;
using System.IO;
using Application.Data;

namespace Application.UI.UserControls
{
    internal class FormFillingManager
    {
        /*-------------------------------------------------------------------------*/

        /**
         * Appelle la méthode de remplissage du fichier excel correspondant au type de formulaire choisi par l'utilisateur
         * 
         * form : Form - L'objet contenant les informations nécessaire à la mise en forme du formulaire
         * parser: Parser - Le parser correspondant au type de fichier à parser
         */
        public void ManageFormFilling(Form form, Parser.Parser parser)
        {
            List<Piece>? data = null;
            
            // Parsing des données
            try
            {
                // Si c'est un formulaire 5 pièces mitutoyo, on récupère les données de tous les fichiers d'un répertoire
                if(form.Type == FormType.FivePieces && parser is TextFileParser)
                {
                    data = this.getDataFromFolder(parser);
                }
                else
                {
                    String fileToParse = this.GetFileToOpen("Sélectionner le fichier à convertir", parser.GetFileExtension());
                    if(fileToParse == "") return;

                    data = parser.ParseFile(fileToParse);
                }
            }
            catch (MeasureTypeNotFoundException e)
            {
                MainWindow.DisplayError(e.Message);
            }
            catch (IncorrectFormatException e)
            {
                MainWindow.DisplayError(e.Message);
            }

            Dictionary<string, string> header = parser.GetHeader();

            if (data == null) return; // à gérer plus tard

            // Remplissage du formulaire
            this.FillForm(form, data, header);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Remplit le formulaire excel
         * 
         * form : Form - L'objet contenant les informations nécessaire à la mise en forme du formulaire
         * data : List<Piece> - Les données à insérer dans le formulaire
         * header : Dictionary<String, String> - Les informations à insérer dans l'entête du formulaire
         */
        public void FillForm(Form form, List<Piece> data, Dictionary<String, String> header)
        {
            try
            {
                String fileToSave = "";

                // Récupération de l'emplacement du formulaire à créer
                fileToSave = this.GetFileToSave();
                if (fileToSave == "") return;

                // Écriture du formulaire
                ExcelWriter writer;

                if(form.Type == FormType.OnePiece) writer = new OnePieceWriter(fileToSave, form);
                else writer = new FivePiecesWriter(fileToSave, form);

                writer.WriteHeader(header, form.DesignLine);
                writer.WriteData(data);
            }
            catch (ExcelFileAlreadyInUseException e)
            {
                MainWindow.DisplayError(e.Message);
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

        /*-------------------------------------------------------------------------*/
    }
}
