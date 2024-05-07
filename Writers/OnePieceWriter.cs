using Excel = Microsoft.Office.Interop.Excel;
using Application.Data;

namespace Application.Writers
{
    internal class OnePieceWriter : ExcelWriter
    {
        private const int MAX_LINES = 22;

        /*-------------------------------------------------------------------------*/

        public OnePieceWriter(string fileName, Form form) : base(fileName, form)
        {
            if(form.Modify)
            {
                this.EraseData(form.FirstLine);
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Crée suffisamment de pages Excel pour écrire les données de la pièce.
         * 
         * La première feuille est la feuille "Mesures" qui contient les données de la pièce.
         * 
         * Si le nombre de lignes à écrire est supérieur à MAX_LINES, des copies de la feuille "Mesures" sont créées.
         */
        public override void CreateWorkSheets()
        {
            int linesToWrite = pieces[0].GetLinesToWriteNumber();

            int pageNumber = linesToWrite / MAX_LINES + 1;

            for (int i = 2; i <= pageNumber; i++)
            {
                workbook.Sheets["Mesures"].Copy(Type.Missing, workbook.Sheets[workbook.Sheets.Count]);
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Écrit les valeurs de mesure des pièces dans les feuilles Excel.
         * 
         */
        public override void WritePiecesValues()
        {
            Excel.Worksheet ws = base.workbook.Sheets["Mesures"];

            List<String> measurePlans = pieces[0].GetMeasurePlans();
            List<List<Data.Data>> pieceData = pieces[0].GetData();

            int linesWritten = 0;
            int pageNumber = 1;

            for (int i = 0; i < pieceData.Count; i++)
            {
                // Écriture du plan
                if (measurePlans[i] != "")
                {
                    ws.Cells[base.currentLine, base.currentColumn + 1].Value = measurePlans[i];
                    base.currentLine++;
                    linesWritten++;
                }

                // Changement de page si l'actuelle est complète
                if (linesWritten == MAX_LINES)
                {
                    pageNumber++;

                    ws = this.workbook.Sheets["Mesures (" + pageNumber.ToString() + ")"];

                    base.currentLine -= linesWritten;
                    linesWritten = 0;
                }

                // Écriture des données ligne par ligne
                for (int j = 0; j < pieceData[i].Count; j++)
                {
                    ws.Cells[base.currentLine, base.currentColumn + 1].Value = pieceData[i][j].Symbol;
                    ws.Cells[base.currentLine, base.currentColumn + 2].Value = pieceData[i][j].NominalValue;
                    ws.Cells[base.currentLine, base.currentColumn + 4].Value = pieceData[i][j].TolerancePlus;
                    ws.Cells[base.currentLine, base.currentColumn + 5].Value = pieceData[i][j].ToleranceMinus;
                    ws.Cells[base.currentLine, base.currentColumn + 6].Value = pieceData[i][j].Value;

                    base.currentLine++;
                    linesWritten++;

                    if (linesWritten == MAX_LINES)
                    {
                        pageNumber++;

                        ws = this.workbook.Sheets["Mesures (" + pageNumber.ToString() + ")"];

                        base.currentLine -= linesWritten;
                        linesWritten = 0;
                    }
                }
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Efface les mesures des pages Excel.
         * 
         * Supprime toutes les pages dont le nom contient Mesures sauf la page Mesures.
         * Supprime les mesures de la première page de mesures.
         */
        public override void EraseData(int firstLine)
        {
            this.excelApp.DisplayAlerts = false;

            // Supprimer toutes les pages dont le nom contient Mesures sauf la page Mesures
            for (int i = workbook.Sheets.Count; i >= 1; i--)
            {
                Excel.Worksheet sheet = (Excel.Worksheet)workbook.Sheets[i];

                if (sheet.Name.Contains("Mesures") && sheet.Name != "Mesures")
                {
                    workbook.Sheets[i].Delete();
                }
            }

            this.excelApp.DisplayAlerts = true;

            // Supprimer les mesures de la première page de mesures
            String start = "B" + firstLine.ToString();
            String end = "I" + (firstLine + MAX_LINES).ToString();
            Excel.Worksheet measuresSheet = (Excel.Worksheet)workbook.Sheets["Mesures"];
            Excel.Range rangeToDelete = measuresSheet.Range[start + ":" + end];
            rangeToDelete.ClearContents();
        }

        /*-------------------------------------------------------------------------*/
    }
}
