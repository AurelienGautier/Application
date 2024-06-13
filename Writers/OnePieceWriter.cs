using Excel = Microsoft.Office.Interop.Excel;
using Application.Data;
using Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    /// <summary>
    /// Represents a writer for a one piece report that writes data to an Excel file.
    /// </summary>
    internal class OnePieceWriter : ExcelWriter
    {
        private const int MAX_LINES = 22;
        private int linesWritten;
        private int pageNumber;

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Initializes a new instance of the OnePieceWriter class with the specified file name and form.
        /// </summary>
        /// <param name="fileName">The name of the Excel file.</param>
        /// <param name="form">The form associated with the writer.</param>
        public OnePieceWriter(string fileName, Form form) : base(fileName, form)
        {
            this.linesWritten = 0;
            this.pageNumber = 1;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Creates enough Excel worksheets to write the piece data.
        /// The first worksheet is the "Mesures" worksheet that contains the piece data.
        /// If the number of lines to write is greater than MAX_LINES, copies of the "Mesures" worksheet are created.
        /// </summary>
        public override void CreateWorkSheets()
        {
            int linesToWrite = pieces[0].GetLinesToWriteNumber();

            int numberOfPages = linesToWrite / MAX_LINES;

            for (int i = 2; i < numberOfPages; i++)
            {
                excelApiLink.CopyWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"], ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + i.ToString() + ")");
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes the measurement values of the pieces to the Excel worksheets.
        /// </summary>
        public override void WritePiecesValues()
        {
            excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"]);

            List<String> measurePlans = pieces[0].GetMeasurePlans();
            List<List<Data.Measure>> pieceData = pieces[0].GetData();


            for (int i = 0; i < pieceData.Count; i++)
            {
                // Writing the plan
                if (measurePlans[i] != "")
                {
                    if (this.isLastLine(pieceData, i, 0) && this.isNextLineEmpty())
                        this.throwIncoherentValueException();

                    excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 1, measurePlans[i]);
                    base.currentLine++;
                    this.linesWritten++;
                }

                // Changing page if the current one is full
                if (this.linesWritten == MAX_LINES)
                {
                    this.ChangePage();
                }

                // Writing the data line by line
                for (int j = 0; j < pieceData[i].Count; j++)
                {
                    if (!base.form.Modify)
                    {
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 1, pieceData[i][j].MeasureType.Symbol);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 2, pieceData[i][j].NominalValue);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 4, pieceData[i][j].TolerancePlus);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 5, pieceData[i][j].ToleranceMinus);
                    }

                    // Throws an error if the number of measures in the report is different from the number of measures in the source file
                    if (form.Modify)
                    {
                        if (excelApiLink.ReadCell(form.Path, base.currentLine, base.currentColumn + 2) == "")
                            this.throwIncoherentValueException();
                        else if (this.isLastLine(pieceData, i, j) && this.isNextLineEmpty())
                            this.throwIncoherentValueException();
                    }

                    excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 6, pieceData[i][j].Value);

                    base.currentLine++;
                    this.linesWritten++;

                    if (this.linesWritten == MAX_LINES)
                    {
                        this.ChangePage();
                    }
                }
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Changes the current page to the next page in the Excel workbook.
        /// </summary>
        private void ChangePage()
        {
            pageNumber++;

            try
            {
                excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + pageNumber.ToString() + ")");
            }
            catch
            {
                excelApiLink.CloseWorkBook(form.Path);

                throw new Exceptions.IncoherentValueException("The number of measures is different between the report to modify and the source file(s).");
            }

            base.currentLine -= linesWritten;
            linesWritten = 0;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Checks if the next line in the Excel worksheet is empty.
        /// </summary>
        /// <returns>True if the next line is empty, false otherwise.</returns>
        private bool isNextLineEmpty()
        {
            if (excelApiLink.ReadCell(form.Path, base.currentLine + 1, base.currentColumn + 2) != "")
                return true;

            return false;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Checks if the current line is the last line of the form to modify.
        /// </summary>
        /// <param name="pieceData">The piece data.</param>
        /// <param name="i">The current index of the measure plan.</param>
        /// <param name="j">The current index of the measurement data within the measure plan.</param>
        /// <returns>True if the current line is the last line, false otherwise.</returns>
        private bool isLastLine(List<List<Data.Measure>> pieceData, int i, int j)
        {
            if (i != pieceData.Count - 1) return false;

            if (pieceData[i].Count == 0 || j == pieceData[i].Count - 1)
                return true;

            return false;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Throws an exception indicating that the number of measures is different between the report to modify and the source file(s).
        /// </summary>
        private void throwIncoherentValueException()
        {
            excelApiLink.CloseWorkBook(form.Path);

            throw new Exceptions.IncoherentValueException("The number of measures is different between the report to modify and the source file(s).");
        }

        /*-------------------------------------------------------------------------*/
    }
}
