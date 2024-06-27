using Application.Data;

namespace Application.Writers
{
    /// <summary>
    /// Represents a writer for a specific type of Excel file that handles five pieces of data.
    /// </summary>
    /// <remarks>
    /// Initializes a new instance of the <see cref="FivePiecesWriter"/> class.
    /// </remarks>
    /// <param name="fileName">The name of the file to be saved.</param>
    /// <param name="form">The form associated with the writer.</param>
    internal class FivePiecesWriter(Form form) : ExcelWriter(form)
    {
        private int pageNumber = 1;
        private int linesWrittenOnCurrentPage = 0;
        private int min = 0;
        private int max = 5;
        private const int MAX_LINES_PER_PAGE = 23;

        /*-------------------------------------------------------------------------*/

        protected override int CalculateNumberOfMeasurePagesToWrite()
        {
            int measureLinesToWrite = pieces[0].GetLinesToWriteNumber();

            int numberOfMeasurePagesToWrite = pieces[0].GetLinesToWriteNumber() / MAX_LINES_PER_PAGE;
            if (measureLinesToWrite % MAX_LINES_PER_PAGE != 0) numberOfMeasurePagesToWrite++;

            int iterations = base.pieces.Count / 5;
            if (base.pieces.Count % 5 != 0) iterations++;

            numberOfMeasurePagesToWrite *= iterations;

            return numberOfMeasurePagesToWrite;
        }

        protected override string GetPageToCopyName(int index)
        {
            int numberOfExistingMeasurePages = base.GetMeasurePagesNumber();

            string nbToCopy = " (" + (index - numberOfExistingMeasurePages + 1).ToString() + ")";
            if (nbToCopy == " (1)") nbToCopy = "";

            string pageToCopyName = ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + nbToCopy;

            return pageToCopyName;
        }

        protected override string GetCopiedPageName(int index)
        {
            return ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + (index + 1).ToString() + ")";
        }

        protected override int GetDataPagesNumber()
        {
            return base.GetMeasurePagesNumber();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes the measurement values of the pieces to the Excel file.
        /// </summary>
        public override void WritePiecesValues()
        {
            excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"]);

            this.max = base.pieces.Count < 5 ? base.pieces.Count : 5;

            int iterations = base.pieces.Count / 5;
            if (base.pieces.Count % 5 != 0) iterations++;

            for (int i = 0; i < iterations; i++)
            {
                this.write5pieces();

                this.min += 5;

                this.max = i == base.pieces.Count / 5 - 1 && base.pieces.Count % 5 != 0 ? base.pieces.Count : this.max + 5;

                if (i < iterations - 1) this.ChangePage();
                this.linesWrittenOnCurrentPage = 0;
                this.currentLine = form.FirstLine;
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes all the values for a group of five pieces to the Excel file.
        /// </summary>
        private void write5pieces()
        {
            List<MeasurePlan> measurePlans = pieces[0].GetMeasurePlans();

            // For each plan
            for (int i = 0; i < measurePlans.Count; i++)
            {
                // Write the plan
                if (measurePlans[i].GetName() != "")
                {
                    base.WriteCell(base.currentLine, base.currentColumn, measurePlans[i].GetName());
                    this.goToNextLine();
                }

                List<Measure> measures = measurePlans[i].GetMeasures();

                // For each measurement in the plan
                for (int j = 0; j < measures.Count; j++)
                {
                    writeMeasure(measures[j]);

                    int columnToWriteValue = base.currentColumn + 3;

                    // Write the values of the pieces
                    for (int k = this.min; k < this.max; k++)
                    {
                        double currentValueToWrite = base.pieces[k].GetMeasurePlans()[i].GetMeasures()[j].Value;

                        columnToWriteValue += 3;
                        base.WriteCell(base.currentLine, columnToWriteValue, currentValueToWrite);
                    }

                    this.goToNextLine();
                }
            }
        }

        /*-------------------------------------------------------------------------*/

        private void clearCurrentLineValues()
        {
            int tempCurrentColumn = 4;

            for (int l = 0; l < 5; l++)
            {
                tempCurrentColumn += 3;
                base.WriteCell(base.currentLine, tempCurrentColumn, "");
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes a measure on the line of the Excel worksheet.
        /// </summary>
        /// <param name="measure">The measure to write.</param>
        private void writeMeasure(Measure measure)
        {
            // Write the measure type, nominal value and tolerances
            if (!base.form.Modify)
            {
                base.WriteCell(base.currentLine, base.currentColumn, measure.MeasureType.Symbol);
                base.WriteCell(base.currentLine, base.currentColumn + 1, measure.NominalValue);
                base.WriteCell(base.currentLine, base.currentColumn + 2, measure.TolerancePlus);
                base.WriteCell(base.currentLine, base.currentColumn + 3, measure.ToleranceMinus);
            }

            if (base.form.Modify) this.clearCurrentLineValues();
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
        /// Switches to the next measurement page.
        /// </summary>
        public void ChangePage()
        {
            this.pageNumber++;

            try
            {
                excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + this.pageNumber.ToString() + ")");
            }
            catch
            {
                base.ThrowIncoherentValueException();
            }

            int col = 7;

            for (int i = this.min; i < this.min + 5; i++)
            {
                excelApiLink.WriteCell(form.Path, 15, col, (i + 1).ToString());
                col += 3;
            }
        }

        /*-------------------------------------------------------------------------*/
    }
}
