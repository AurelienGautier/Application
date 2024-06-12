﻿using Excel = Microsoft.Office.Interop.Excel;
using Application.Data;

namespace Application.Writers
{
    internal class OnePieceWriter : ExcelWriter
    {
        private const int MAX_LINES = 22;
        private int linesWritten;
        private int pageNumber;

        /*-------------------------------------------------------------------------*/

        public OnePieceWriter(string fileName, Form form) : base(fileName, form)
        {
            this.linesWritten = 0;
            this.pageNumber = 1;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Crée suffisamment de pages Excel pour écrire les données de la pièce.
         * 
         * La première feuille est la feuille "Mesures" qui contient les données de la pièce.
         * 
         * Si le nombre de lignes à écrire est supérieur à MAX_LINES, des copies de la feuille "Mesures" sont créées.
         * 
         */
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

        /**
         * Écrit les valeurs de mesure des pièces dans les feuilles Excel.
         * 
         */
        public override void WritePiecesValues()
        {
            excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"]);

            List<String> measurePlans = pieces[0].GetMeasurePlans();
            List<List<Data.Measure>> pieceData = pieces[0].GetData();

            for (int i = 0; i < pieceData.Count; i++)
            {
                // Écriture du plan
                if (measurePlans[i] != "")
                {
                    excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 1, measurePlans[i]);
                    base.currentLine++;
                    this.linesWritten++;
                }

                // Changement de page si l'actuelle est complète
                if (this.linesWritten == MAX_LINES)
                {
                    this.ChangePage();
                }

                // Écriture des données ligne par ligne
                for (int j = 0; j < pieceData[i].Count; j++)
                {
                    if(!base.form.Modify)
                    {
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 1, pieceData[i][j].MeasureType.Symbol);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 2, pieceData[i][j].NominalValue);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 4, pieceData[i][j].TolerancePlus);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 5, pieceData[i][j].ToleranceMinus);
                    }

                    if(base.form.Modify && excelApiLink.ReadCell(form.Path, base.currentLine, base.currentColumn + 6) != "") 
                    {
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 6, pieceData[i][j].Value);
                    }

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

        private void ChangePage()
        {
            pageNumber++;

            try
            {
                excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + pageNumber.ToString() + ")");
            }
            catch
            {
                if (form.Modify) return;
            }

            base.currentLine -= linesWritten;
            linesWritten = 0;
        }
    }
}
