﻿using Application.Data;

namespace Application.Writers
{
    /// <summary>
    /// Represents a writer for a specific type of Excel file that handles five pieces of data.
    /// </summary>
    internal class FivePiecesWriter : ExcelWriter
    {
        private int pageNumber;
        private readonly List<List<String>> measurePlans;
        private readonly List<List<List<Data.Data>>> pieceData;
        private int linesWritten;
        private int min;
        private int max;
        private const int MAX_LINES = 23;

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Initializes a new instance of the <see cref="FivePiecesWriter"/> class.
        /// </summary>
        /// <param name="fileName">The name of the file to be saved.</param>
        /// <param name="form">The form associated with the writer.</param>
        public FivePiecesWriter(string fileName, Form form) : base(fileName, form)
        {
            this.pageNumber = 1;
            this.measurePlans = new List<List<String>>();
            this.pieceData = new List<List<List<Data.Data>>>();
            this.linesWritten = 0;
            this.min = 0;
            this.max = 5;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Creates all the necessary Excel worksheets to insert all the data.
        /// </summary>
        public override void CreateWorkSheets()
        {
            int totalPageNumber = pieces[0].GetLinesToWriteNumber() / MAX_LINES + 1;

            int iterations = base.pieces.Count / 5;
            if (base.pieces.Count % 5 != 0) iterations++;

            totalPageNumber *= iterations;

            for (int i = 2; i <= totalPageNumber; i++)
            {
                excelApiLink.CopyWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"], ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + i.ToString() + ")");
            }
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
                this.pieceData.Add(base.pieces[i].GetData());
            }

            int iterations = pieceData.Count / 5;
            if (pieceData.Count % 5 != 0) iterations++;

            for (int i = 0; i < iterations; i++)
            {
                this.write5pieces();

                this.min += 5;

                if (i == pieceData.Count / 5 - 1 && pieceData.Count % 5 != 0) this.max = pieceData.Count;
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
            for (int i = 0; i < pieceData[0].Count; i++)
            {
                // Write the plan
                if (measurePlans[0][i] != "")
                {
                    excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn, measurePlans[0][i]);
                    base.currentLine++;
                    this.linesWritten++;
                }

                // Change page if the current one is full
                if (this.linesWritten == MAX_LINES) { this.ChangePage(); }

                // For each measurement in the plan
                for (int j = 0; j < pieceData[0][i].Count; j++)
                {
                    if (!base.form.Modify)
                    {
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn, pieceData[0][i][j].Symbol);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 1, pieceData[0][i][j].NominalValue);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 2, pieceData[0][i][j].TolerancePlus);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 3, pieceData[0][i][j].ToleranceMinus);
                    }

                    base.currentColumn += 3;

                    // Write the values of the pieces
                    for (int k = this.min; k < this.max; k++)
                    {
                        base.currentColumn += 3;
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn, pieceData[k][i][j].Value);
                    }

                    base.currentColumn -= (3 + 3 * (this.max - this.min));

                    base.currentLine++;
                    this.linesWritten++;

                    // Change page if the current one is full
                    if (this.linesWritten == MAX_LINES) { this.ChangePage(); }
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
            excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + this.pageNumber.ToString() + ")");

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
    }
}
