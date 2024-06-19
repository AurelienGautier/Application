using Application.Data;

namespace Application.Writers
{
    /// <summary>
    /// Represents a writer for a specific type of Excel file that handles five pieces of data.
    /// </summary>
    internal class FivePiecesWriter : ExcelWriter
    {
        private int pageNumber;
        private int linesWritten;
        readonly private List<List<MeasurePlan>> measurePlans;
        private int min;
        private int max;
        private const int MAX_LINES_PER_PAGE = 23;

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Initializes a new instance of the <see cref="FivePiecesWriter"/> class.
        /// </summary>
        /// <param name="fileName">The name of the file to be saved.</param>
        /// <param name="form">The form associated with the writer.</param>
        public FivePiecesWriter(string fileName, Form form) : base(fileName, form)
        {
            this.pageNumber = 1;
            this.measurePlans = new List<List<MeasurePlan>>();
            this.linesWritten = 0;
            this.min = 0;
            this.max = 5;
        }

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
            int numberOfExistingMeasurePages = base.getMeasurePagesNumber();

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
            return base.getMeasurePagesNumber();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes the measurement values of the pieces to the Excel file.
        /// </summary>
        public override void WritePiecesValues()
        {
            excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"]);

            for (int i = 0; i < base.pieces.Count; i++)
            {
                this.measurePlans.Add(base.pieces[i].GetMeasurePlans());
            }

            this.max = this.measurePlans.Count < 5 ? this.measurePlans.Count : 5;

            int iterations = this.measurePlans.Count / 5;
            if (this.measurePlans.Count % 5 != 0) iterations++;

            for (int i = 0; i < iterations; i++)
            {
                this.write5pieces();

                this.min += 5;

                if (i == this.measurePlans.Count / 5 - 1 && this.measurePlans.Count % 5 != 0) this.max = this.measurePlans.Count;
                else this.max += 5;

                if (i < iterations - 1) this.ChangePage();
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes all the values for a group of five pieces to the Excel file.
        /// </summary>
        private void write5pieces()
        {
            // For each plan
            for (int i = 0; i < measurePlans[0].Count; i++)
            {
                // Write the plan
                if (measurePlans[0][i].GetName() != "")
                {
                    excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn, measurePlans[0][i].GetName());
                    base.currentLine++;
                    this.linesWritten++;
                }

                // Change page if the current one is full
                if (this.linesWritten == MAX_LINES_PER_PAGE) { this.ChangePage(); }

                List<Measure> measures = measurePlans[0][i].GetMeasures();

                // For each measurement in the plan
                for (int j = 0; j < measures.Count; j++)
                {
                    if (!base.form.Modify)
                    {
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn, measures[j].MeasureType.Symbol);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 1, measures[j].NominalValue);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 2, measures[j].TolerancePlus);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 3, measures[j].ToleranceMinus);
                    }

                    base.currentColumn += 3;

                    if (base.form.Modify)
                    {
                        int tempCurrentColumn = 4;

                        for(int l = 0; l < 5; l++)
                        {
                            tempCurrentColumn += 3;
                            excelApiLink.WriteCell(form.Path, base.currentLine, tempCurrentColumn, "");
                        }
                    }

                    // Write the values of the pieces
                    for (int k = this.min; k < this.max; k++)
                    {
                        base.currentColumn += 3;
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn, measurePlans[k][i].GetMeasures()[j].Value);
                    }

                    base.currentColumn -= (3 + 3 * (this.max - this.min));

                    base.currentLine++;
                    this.linesWritten++;

                    // Change page if the current one is full
                    if (this.linesWritten == MAX_LINES_PER_PAGE) { this.ChangePage(); }
                }
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
                Console.WriteLine(base.getMeasurePagesNumber());
                Console.WriteLine(this.pageNumber);
                base.throwIncoherentValueException();
            }

            base.currentLine = 17;
            this.linesWritten = 0;

            int col = 7;

            for (int i = this.min; i < this.min + 5; i++)
            {
                excelApiLink.WriteCell(form.Path, 15, col, (i + 1).ToString());
                col += 3;
            }
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
        /// Checks if the next line in the Excel worksheet is empty.
        /// </summary>
        /// <returns>True if the next line is empty, false otherwise.</returns>
        private bool isNextLineEmpty()
        {
            if (excelApiLink.ReadCell(form.Path, base.currentLine + 1, base.currentColumn) != ""
                || excelApiLink.ReadCell(form.Path, base.currentLine + 1, base.currentColumn + 1) != "")
                return false;

            return true;
        }

        /*-------------------------------------------------------------------------*/
    }
}
