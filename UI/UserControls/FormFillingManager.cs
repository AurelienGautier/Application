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
        public void ManageFormFilling(Form form, Parser.Parser parser, List<Standard> standards, String dataPath, String fileToSave)
        {
            // Parsing the data
            List<Piece>? data = this.GetData(form.DataFrom, parser, dataPath);
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
        private List<Piece>? GetData(DataFrom dataFrom, Parser.Parser parser, String dataPath)
        {
            List<Piece>? data;

            // Parsing the data
            try
            {
                // If it's a 5-piece Mitutoyo form, retrieve data from all files in a directory
                if (dataFrom == DataFrom.Folder)
                {
                    data = this.GetDataFromFolder(parser, dataPath);
                }
                else
                {
                    data = parser.ParseFile(dataPath);
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
        /// Retrieves the data from all files in a directory.
        /// </summary>
        /// <param name="parser">The parser corresponding to the file type to parse.</param>
        /// <param name="folderPath">The path to the directory.</param>
        /// <returns>The list of data pieces retrieved from the files in the directory.</returns>
        private List<Data.Piece>? GetDataFromFolder(Parser.Parser parser, String folderPath)
        {
            DirectoryInfo directory = new DirectoryInfo(folderPath);

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
                        MainWindow.DisplayError("The file format of " + fileName + " is incorrect.");
                        return Enumerable.Empty<Data.Piece>();
                    }
                    catch (MeasureTypeNotFoundException)
                    {
                        MainWindow.DisplayError("A measure type was not found in the file " + fileName);
                        return Enumerable.Empty<Data.Piece>();
                    }
                })
                .ToList();

            return data;
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
