using Application.Exceptions;
using Application.Writers;
using Microsoft.Win32;
using System.IO;
using Application.Data;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Manages the filling of forms in an Excel file based on user-selected form type.
    /// </summary>
    internal class FormFillingManager
    {
        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Calls the method to fill the corresponding Excel file based on the user-selected form type.
        /// </summary>
        /// <param name="form">The object containing the necessary information for formatting the form.</param>
        /// <param name="parser">The parser corresponding to the file type to parse.</param>
        /// <param name="standards">The list of standards to be used for filling the form.</param>
        /// <param name="dataPath">The path to the data file.</param>
        /// <param name="fileToSave">The path to save the filled form.</param>
        public void ManageFormFilling(Form form, Parser.Parser parser, List<Standard> standards, String fileToSave)
        {
            // Parsing the data
            List<Piece>? data = this.GetData(form, parser);
            if (data == null) return;

            // Filling the form
            this.FillForm(form, data, standards, fileToSave);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Retrieves the data to be inserted into the form based on the form's data source type.
        /// </summary>
        /// <param name="dataFrom">The data source type of the form.</param>
        /// <param name="parser">The parser corresponding to the file type to parse.</param>
        /// <param name="dataPath">The path to the data file.</param>
        /// <returns>The list of data pieces to be inserted into the form.</returns>
        private List<Piece>? GetData(Form form, Parser.Parser parser)
        {
            List<Piece> data = new List<Piece>();
            int? measureNumber = null;

            // Parsing the data
            try
            {
                for (int i = 0; i < form.SourceFiles.Count; i++)
                {
                    List<Piece> newPieces = parser.ParseFile(form.SourceFiles[i]);

                    Console.WriteLine(form.SourceFiles[i]);

                    if(measureNumber == null)
                    {
                        measureNumber = newPieces[0].GetLinesToWriteNumber();
                    }
                    else if (measureNumber != newPieces[0].GetLinesToWriteNumber())
                    {
                        MainWindow.DisplayError("Le nombre de mesures des pièces n'est pas le même entre le fichier numéro 1 et le fichier numéro " + (i + 1));
                        return null;
                    }

                    data.AddRange(parser.ParseFile(form.SourceFiles[i]));
                }
            }
            catch (MeasureTypeNotFoundException e)
            {
                MainWindow.DisplayError(e.Message);
                return null;
            }
            catch (IncorrectFormatException e)
            {
                MainWindow.DisplayError(e.Message);
                return null;
            }

            return data;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Fills the Excel form.
        /// </summary>
        /// <param name="form">The object containing the necessary information for formatting the form.</param>
        /// <param name="data">The data to be inserted into the form.</param>
        /// <param name="standards">The information to be inserted into the form header.</param>
        /// <param name="fileToSave">The path to save the filled form.</param>
        public void FillForm(Form form, List<Piece> data, List<Standard> standards, String fileToSave)
        {
            try
            {
                // Writing the form
                ExcelWriter writer;

                if (form.Type == FormType.OnePiece) writer = new OnePieceWriter(fileToSave, form);
                else if (form.Type == FormType.FivePieces) writer = new FivePiecesWriter(fileToSave, form);
                else writer = new CapabilityWriter(fileToSave, form);

                writer.WriteData(data, standards);
            }
            catch (ExcelFileAlreadyInUseException e)
            {
                MainWindow.DisplayError(e.Message);
            }
            catch (ConfigDataException e)
            {
                MainWindow.DisplayError(e.Message);
            }
            catch (IncoherentValueException e)
            {
                MainWindow.DisplayError(e.Message);
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Opens a dialog window to select a file path.
        /// </summary>
        /// <param name="title">The title of the dialog window.</param>
        /// <param name="extensions">The allowed file extensions.</param>
        /// <returns>The selected file path.</returns>
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

        /// <summary>
        /// Opens a dialog window to select one or multiple file path.
        /// </summary>
        /// <param name="title">The title of the dialog window.</param>
        /// <param name="extensions">The allowed file extensions.</param>
        /// <param name="multiSelect">True if multiple files can be selected, false otherwise.</param>
        /// <returns>The selected file path.</returns>
        public List<String> GetFilesToOpen(String title, String extensions, bool multiSelect)
        {
            var dialog = new OpenFileDialog();
            dialog.Title = title;
            dialog.Filter = extensions;
            dialog.Multiselect = multiSelect;

            List<String> fileNames = new List<String>();

            if (dialog.ShowDialog() == true)
            {
                fileNames.AddRange(dialog.FileNames);
            }

            return fileNames;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Opens a dialog window to select a folder path.
        /// </summary>
        /// <param name="title">The title of the dialog window.</param>
        /// <returns>The selected folder path.</returns>
        public String GetFolderToOpen(String title)
        {
            var dialog = new OpenFolderDialog();
            dialog.Title = title;

            String folderName = "";

            if (dialog.ShowDialog() == true) folderName = dialog.FolderName;

            return folderName;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Opens a dialog window to select the path where to save a file.
        /// </summary>
        /// <returns>The selected file path.</returns>
        public String GetFileToSave()
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "(*.xlsx;*.xlsm)|*.xlsx;*.xlsm";

            saveFileDialog.FileName = "rapport.xlsm";

            String fileName = "";

            if (saveFileDialog.ShowDialog() == true)
                fileName = saveFileDialog.FileName;

            // Remove the extension from the Excel file
            if (fileName.Length > 5)
                fileName = fileName.Remove(fileName.Length - 5);

            return fileName;
        }

        /*-------------------------------------------------------------------------*/
    }
}
