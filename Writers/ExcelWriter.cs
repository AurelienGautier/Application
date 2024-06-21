using Application.Data;
using Application.Exceptions;
using Application.Facade;

namespace Application.Writers
{
    /// <summary>
    /// Base class for writing data to an Excel file.
    /// </summary>
    internal abstract class ExcelWriter
    {
        private readonly string fileToSavePath;
        protected int currentLine;
        protected int currentColumn;
        protected List<Data.Piece> pieces;
        protected Form form;
        protected ExcelLibraryLinkSingleton excelApiLink;

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Constructor for the ExcelWriter class.
        /// </summary>
        /// <param name="fileName">The path of the file to save.</param>
        /// <param name="form">The form object containing the initial line and column values.</param>
        protected ExcelWriter(string fileName, Form form)
        {
            this.fileToSavePath = fileName;
            this.currentLine = form.FirstLine;
            this.currentColumn = form.FirstColumn;
            this.form = form;

            ExcelLibraryLinkSingleton.Instance.OpenWorkBook(form.Path);

            this.pieces = [];

            this.excelApiLink = ExcelLibraryLinkSingleton.Instance;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes the data of the pieces to the Excel file.
        /// </summary>
        /// <param name="data">The list of pieces to write.</param>
        /// <param name="standards">The list of standards to write.</param>
        public void WriteData(List<Data.Piece> data)
        {
            this.pieces = data;

            writeHeader(data[0].GetHeader());

            CreateWorkSheets();

            WritePiecesValues();

            if (this.form.Sign)
            {
                signForm();

                exportToPdf();
            }

            saveAndQuit();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes the header of the Excel report.
        /// </summary>
        /// <param name="header">The dictionary containing the header information.</param>
        /// <param name="standards">The list of standards to write.</param>
        private void writeHeader(Header header)
        {
            if (!form.Modify)
            {
                excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["HeaderPage"]);

                excelApiLink.WriteCell(form.Path, form.DesignLine, 4, header.Designation);
                excelApiLink.WriteCell(form.Path, form.DesignLine + 2, 4, header.PlanNb);
                excelApiLink.WriteCell(form.Path, form.DesignLine + 4, 4, header.Index);
                excelApiLink.WriteCell(form.Path, 14, 1, "N° " + header.ObservationNum);
                excelApiLink.WriteCell(form.Path, 38, 8, header.PieceReceptionDate);
                excelApiLink.WriteCell(form.Path, 40, 4, header.Observations);

                this.writeClient(header.ClientName);
                this.writeStandards(this.form.Standards);
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes the client information in the header of the form by retrieving it from the client file.
        /// </summary>
        /// <param name="client">The name of the client.</param>
        private void writeClient(string client)
        {
            string clientWorkbookPath = Environment.CurrentDirectory + "\\res\\ADRESSE";
            excelApiLink.OpenWorkBook(clientWorkbookPath);
            excelApiLink.ChangeWorkSheet(clientWorkbookPath, "ADRESSE");

            int currentLineWs2 = 2;

            // While the current line is not empty and the client has not been found
            while (excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 2) != ""
                && excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 2) != client)
            {
                currentLineWs2++;
            }

            string address = "";
            string bp = "";
            string postalCode = "";
            string city = "";

            if (excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 2) != "")
            {
                address = excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 3);
                bp = excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 4);
                postalCode = excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 5);
                city = excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 6);
            }

            excelApiLink.CloseWorkBook(clientWorkbookPath);

            excelApiLink.WriteCell(form.Path, form.ClientLine, 4, client);
            excelApiLink.WriteCell(form.Path, form.ClientLine + 1, 4, address);
            excelApiLink.WriteCell(form.Path, form.ClientLine + 2, 4, bp);
            excelApiLink.WriteCell(form.Path, form.ClientLine + 3, 4, postalCode);
            excelApiLink.WriteCell(form.Path, form.ClientLine + 3, 5, city);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes the standards information in the header page of the form.
        /// </summary>
        /// <param name="standards">The list of standards to write.</param>
        private void writeStandards(List<Standard> standards)
        {
            int linesToShift = standards.Count * 4;

            // Shift the values downwards
            excelApiLink.ShiftLines(form.Path, form.StandardLine, 1, 15, linesToShift);

            int standardLine = form.StandardLine;

            foreach (Standard standard in standards)
            {
                excelApiLink.WriteCell(form.Path, standardLine, 1, "Moyen:");
                excelApiLink.WriteCell(form.Path, standardLine, 4, standard.Name);

                excelApiLink.WriteCell(form.Path, standardLine + 1, 1, "Raccordement:");
                excelApiLink.WriteCell(form.Path, standardLine + 1, 4, standard.Raccordement);

                excelApiLink.MergeCells(form.Path, standardLine + 2, 4, standardLine + 2, 5);

                excelApiLink.WriteCell(form.Path, standardLine + 2, 1, "Validité:");
                excelApiLink.WriteCell(form.Path, standardLine + 2, 4, standard.Validity);

                standardLine += 4;
            }
        }

        /*-------------------------------------------------------------------------*/

        protected abstract int CalculateNumberOfMeasurePagesToWrite();
        protected abstract string GetPageToCopyName(int index);
        protected abstract string GetCopiedPageName(int index);

        /// <summary>
        /// Creates the necessary worksheets (delegated to child classes).
        /// </summary>
        public void CreateWorkSheets()
        {
            int numberOfMeasurePagesToWrite = CalculateNumberOfMeasurePagesToWrite();
            int numberOfExistingMeasurePages = GetDataPagesNumber();

            // Delete the extra worksheets if the number of existing worksheets is greater than the number of necessary worksheets
            for (int i = numberOfExistingMeasurePages; i > numberOfMeasurePagesToWrite; i--)
            {
                string pageName = ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + i.ToString() + ")";
                this.DeleteWorkSheet(pageName);
            }

            // Creates all the necessary worksheets that don't exist yet
            for (int i = numberOfExistingMeasurePages; i < numberOfMeasurePagesToWrite; i++)
            {
                string pageToCopyName = GetPageToCopyName(i);
                string copiedPageName = GetCopiedPageName(i);
                this.CopyWorkSheet(pageToCopyName, copiedPageName);
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes the values of the pieces to the Excel file (delegated to child classes).
        /// </summary>
        public abstract void WritePiecesValues();

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Pastes the signature image on the form.
        /// </summary>
        private void signForm()
        {
            excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["HeaderPage"]);

            if (ConfigSingleton.Instance.Signature == null)
                throw new ConfigDataException("Il est impossible de signer le formulaire, la signature n'a pas été trouvée ou son format est incorrect");

            excelApiLink.PasteImage(form.Path, form.LineToSign, form.ColumnToSign, ConfigSingleton.Instance.Signature);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Exports the first page of the Excel file to PDF.
        /// </summary>
        private void exportToPdf()
        {
            excelApiLink.ExportToPdf(form.Path, fileToSavePath.Replace(".xlsx", ".pdf"));
            excelApiLink.DeleteImage(form.Path, form.LineToSign, form.ColumnToSign);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Saves the file and closes the application.
        /// </summary>
        private void saveAndQuit()
        {
            excelApiLink.SaveWorkBook(form.Path, fileToSavePath);

            excelApiLink.CloseWorkBook(form.Path);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Throws an exception indicating that the number of measures is different between the report to modify and the source file(s).
        /// </summary>
        protected void ThrowIncoherentValueException()
        {
            excelApiLink.CloseWorkBook(form.Path);
            
            throw new Exceptions.IncoherentValueException("Le nombre de mesures est différent entre le rapport à modifier et le(s) fichier(s) source(s).");
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the number of pages containing measurement values in the Excel file.
        /// </summary>
        /// <returns>The number of pages</returns>
        protected int GetMeasurePagesNumber()
        {
            int pageNumber = 0;

            for (int i = 1; i <= excelApiLink.GetNumberOfPages(form.Path); i++)
            {
                string pageName = excelApiLink.GetPageName(form.Path, i);

                if (pageName.StartsWith(ConfigSingleton.Instance.GetPageNames()["MeasurePage"]))
                    pageNumber++;
            }

            return pageNumber;
        }

        /*-------------------------------------------------------------------------*/

        abstract protected int GetDataPagesNumber();

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Uses the ExcelLibraryLinkSingleton to delete a worksheet.
        /// </summary>
        /// <param name="sheetName">The name of the worksheet to delete</param>
        protected void DeleteWorkSheet(string sheetName)
        {
            excelApiLink.DeleteWorkSheet(form.Path, sheetName);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Uses the ExcelLibraryLinkSingleton to copy a worksheet.
        /// </summary>
        /// <param name="sheetName">The name of the sheet to copy</param>
        /// <param name="newSheetName">The name of the copied sheet</param>
        protected void CopyWorkSheet(string sheetName, string newSheetName)
        {
            excelApiLink.CopyWorkSheet(form.Path, sheetName, newSheetName);
        }

        /*-------------------------------------------------------------------------*/

        protected void WriteCell(int row, int col, string value)
        {
            excelApiLink.WriteCell(form.Path, row, col, value);
        }

        protected void WriteCell(int row, int col, double value)
        {
            excelApiLink.WriteCell(form.Path, row, col, value);
        }

        /*-------------------------------------------------------------------------*/
    }
}