using Application.Data;

namespace Application.Writers
{
    /// <summary>
    /// Represents a writer for a one piece report that writes data to an Excel file.
    /// </summary>
    internal class OnePieceWriter : ExcelWriter
    {
        private const int MAX_LINES_PER_PAGE = 22;
        private int linesWrittenOnCurrentPage;
        private int pageNumber;

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Initializes a new instance of the OnePieceWriter class with the specified file name and form.
        /// </summary>
        /// <param name="fileName">The name of the Excel file.</param>
        /// <param name="form">The form associated with the writer.</param>
        public OnePieceWriter(string fileName, Form form) : base(fileName, form)
        {
            this.linesWrittenOnCurrentPage = 0;
            this.pageNumber = 1;
        }

        /*-------------------------------------------------------------------------*/

        protected override int CalculateNumberOfMeasurePagesToWrite()
        {
            int measureLinesToWrite = pieces[0].GetLinesToWriteNumber();

            int numberOfMeasurePagesToWrite = measureLinesToWrite / MAX_LINES_PER_PAGE;
            if (measureLinesToWrite % MAX_LINES_PER_PAGE != 0) numberOfMeasurePagesToWrite++;

            return numberOfMeasurePagesToWrite;
        }

        protected override string GetPageToCopyName(int index)
        {
            return ConfigSingleton.Instance.GetPageNames()["MeasurePage"];
        }

        protected override string GetCopiedPageName(int index)
        {
            return ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + (index + 1).ToString() + ")";
        }

        protected override int GetDataPagesNumber()
        {
            return base.getMeasurePagesNumber();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes the measurement values of the pieces to the Excel worksheets.
        /// </summary>
        public override void WritePiecesValues()
        {
            // Change the worksheet to the first measure page
            excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"]);

            List<MeasurePlan> measurePlans = pieces[0].GetMeasurePlans();

            // For each measure plan of the piece
            for (int i = 0; i < measurePlans.Count; i++)
            {
                // Writing the name of the measure plan
                if (measurePlans[i].GetName() != "")
                {
                    this.writeCell(base.currentLine, base.currentColumn + 1, measurePlans[i].GetName());
                    this.goToNextLine();
                }

                List<Measure> measures = measurePlans[i].GetMeasures();

                // Writing the data line by line
                for (int j = 0; j < measures.Count; j++)
                {
                    writeMeasure(measures[j]);
                }
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes a measure into the Excel worksheet.
        /// </summary>
        /// <param name="measure">The measure to write.</param>
        private void writeMeasure(Measure measure)
        {
            if (!base.form.Modify)
            {
                this.writeCell(base.currentLine, base.currentColumn + 1, measure.MeasureType.Symbol);
                this.writeCell(base.currentLine, base.currentColumn + 2, measure.NominalValue);
                this.writeCell(base.currentLine, base.currentColumn + 4, measure.TolerancePlus);
                this.writeCell(base.currentLine, base.currentColumn + 5, measure.ToleranceMinus);
            }

            this.writeCell(base.currentLine, base.currentColumn + 6, measure.Value);
            excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 6, measure.Value);

            this.goToNextLine();
        }

        /*-------------------------------------------------------------------------*/

        private void writeCell(int row, int col, string value)
        {
            excelApiLink.WriteCell(form.Path, row, col, value);
        }

        private void writeCell(int row, int col, double value)
        {
            excelApiLink.WriteCell(form.Path, row, col, value);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Moves to the next line in the Excel worksheet.
        /// </summary>
        private void goToNextLine()
        {
            base.currentLine++;
            this.linesWrittenOnCurrentPage++;

            // Change page if the current one is full
            if (this.linesWrittenOnCurrentPage == MAX_LINES_PER_PAGE)
            {
                this.ChangePage();
                this.linesWrittenOnCurrentPage = 0;
                this.currentLine -= MAX_LINES_PER_PAGE;
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

                throw new Exceptions.IncoherentValueException("Le nombre de pages n'est pas suffisant");
            }
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
