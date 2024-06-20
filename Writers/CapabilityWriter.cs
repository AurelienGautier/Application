using Application.Data;

namespace Application.Writers
{
    class CapabilityWriter : ExcelWriter
    {
        private const int MAX_LINES_PER_PAGE = 25;
        private int linesWrittenOnCurrentPage = 0;
        private int pageNumber = 1;

        public CapabilityWriter(String fileName, Form form) : base(fileName, form) { }

        /*-------------------------------------------------------------------------*/

        protected override int CalculateNumberOfMeasurePagesToWrite()
        {
            if (form.CapabilityMeasureNumber == null)
                throw new Exceptions.IncorrectValuesToTreatException("Le nombre de mesures de capacité n'a pas été renseigné.");

            return form.CapabilityMeasureNumber.Count;
        }

        protected override string GetPageToCopyName(int index)
        {
            return "Capa";
        }

        protected override string GetCopiedPageName(int index)
        {
            return "Capa (" + (index + 1) + ")";
        }

        protected override int GetDataPagesNumber()
        {
            return this.getCapaPagesNumber();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes the values of each capability measure in a different worksheet
        /// </summary>
        /// <exception cref="Exceptions.IncoherentValueException"></exception>
        public override void WritePiecesValues()
        {
            excelApiLink.ChangeWorkSheet(form.Path, "Capa");

            if (form.CapabilityMeasureNumber == null) return;
            List<int> capabilityMeasureNumber = form.CapabilityMeasureNumber;

            // Write the values of the pieces in the capability form
            for (int i = 0; i < capabilityMeasureNumber.Count; i++)
            {
                if (i > 0) this.changePage();

                int num = capabilityMeasureNumber[i];
                foreach (Piece piece in pieces)
                {
                    if (num >= piece.GetMeasurePlans()[0].GetMeasures().Count)
                        throw new Exceptions.IncoherentValueException("Le numéro de mesure de capacité " + num + " indiqué est incorrect.");

                    double currentValue = piece.GetMeasurePlans()[0].GetMeasures()[num].Value;

                    excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn, currentValue);

                    this.goToNextLine();
                }
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the number of pages containing measurement values in the Excel file.
        /// </summary>
        /// <returns>The number of pages</returns>
        private int getCapaPagesNumber()
        {
            int capaPageNumber = 0;

            for (int i = 1; i <= excelApiLink.GetNumberOfPages(form.Path); i++)
            {
                string pageName = excelApiLink.GetPageName(form.Path, i);

                if (pageName.StartsWith("Capa"))
                    capaPageNumber++;
            }

            return capaPageNumber;
        }

        /*-------------------------------------------------------------------------*/

        private void goToNextLine()
        {
            base.currentLine++;
            this.linesWrittenOnCurrentPage++;

            // Change column if the current one is full
            if (this.linesWrittenOnCurrentPage == MAX_LINES_PER_PAGE)
            {
                this.linesWrittenOnCurrentPage = 0;
                base.currentLine -= MAX_LINES_PER_PAGE;
                base.currentColumn++;
            }
        }

        /*-------------------------------------------------------------------------*/

        private void changePage()
        {
            this.pageNumber++;
            this.linesWrittenOnCurrentPage = 0;
            base.currentLine = base.form.FirstLine;
            base.currentColumn = base.form.FirstColumn;

            try
            {
                excelApiLink.ChangeWorkSheet(form.Path, "Capa (" + (this.pageNumber) + ")");
            }
            catch
            {
                excelApiLink.DisplayWorkSheets(form.Path);

                excelApiLink.CloseWorkBook(form.Path);

                throw new Exceptions.IncoherentValueException("Le nombre de pages n'est pas suffisant");
            }
        }

        /*-------------------------------------------------------------------------*/
    }
}
