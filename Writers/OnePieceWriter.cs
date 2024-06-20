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
                    base.WriteCell(base.currentLine, base.currentColumn + 1, measurePlans[i].GetName());
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
        /// Writes a measure on the line of the Excel worksheet.
        /// </summary>
        /// <param name="measure">The measure to write.</param>
        private void writeMeasure(Measure measure)
        {
            if (!base.form.Modify)
            {
                base.WriteCell(base.currentLine, base.currentColumn + 1, measure.MeasureType.Symbol);
                base.WriteCell(base.currentLine, base.currentColumn + 2, measure.NominalValue);
                base.WriteCell(base.currentLine, base.currentColumn + 4, measure.TolerancePlus);
                base.WriteCell(base.currentLine, base.currentColumn + 5, measure.ToleranceMinus);
            }

            base.WriteCell(base.currentLine, base.currentColumn + 6, measure.Value);
            excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 6, measure.Value);

            this.goToNextLine();
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
                this.changePage();
                this.linesWrittenOnCurrentPage = 0;
                this.currentLine -= MAX_LINES_PER_PAGE;
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Changes the current page to the next page in the Excel workbook.
        /// </summary>
        private void changePage()
        {
            this.pageNumber++;

            try
            {
                excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + this.pageNumber.ToString() + ")");
            }
            catch
            {
                excelApiLink.CloseWorkBook(form.Path);

                throw new Exceptions.IncoherentValueException("Le nombre de pages n'est pas suffisant");
            }
        }

        /*-------------------------------------------------------------------------*/
    }
}
