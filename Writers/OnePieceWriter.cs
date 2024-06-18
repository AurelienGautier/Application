using Application.Data;

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

            // Pages that already exist in the report don't need to be created again
            int firstPageToCreate = base.form.Modify ? base.getMeasurePagesNumber() + 1 : 2;

            // If the number of pages to create is less than the number of pages in the file, delete the extra pages
            if (base.form.Modify && firstPageToCreate > numberOfPages)
            {
                for (int i = firstPageToCreate - 1; i > numberOfPages; i--)
                {
                    excelApiLink.DeleteWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + i.ToString() + ")");
                }
            }

            for (int i = firstPageToCreate; i < numberOfPages; i++)
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

            List<MeasurePlan> measurePlans = pieces[0].GetMeasurePlans();


            for (int i = 0; i < measurePlans.Count; i++)
            {
                // Writing the plan
                if (measurePlans[i].GetName() != "")
                {
                    excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 1, measurePlans[i].GetName());
                    base.currentLine++;
                    this.linesWritten++;
                }

                // Changing page if the current one is full
                if (this.linesWritten == MAX_LINES)
                {
                    this.ChangePage();
                }

                List<Measure> measures = measurePlans[i].GetMeasures();

                // Writing the data line by line
                for (int j = 0; j < measures.Count; j++)
                {
                    if (!base.form.Modify)
                    {
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 1, measures[j].MeasureType.Symbol);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 2, measures[j].NominalValue);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 4, measures[j].TolerancePlus);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 5, measures[j].ToleranceMinus);
                    }

                    excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 6, measures[j].Value);

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
                return false;

            return true;
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
    }
}
