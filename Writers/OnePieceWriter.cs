using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    internal class OnePieceWriter : ExcelWriter
    {
        private const int MAX_LINES = 22;

        /*-------------------------------------------------------------------------*/

        public OnePieceWriter(string fileName, int firstLine, string formPath, bool modify) : base(fileName, firstLine, 1, formPath, modify)
        {
            if(base.modify)
            {
                this.EraseData(firstLine);
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
                    base.currentColumn++;
                    ws.Cells[base.currentLine, base.currentColumn].Value = measurePlans[i];
                    base.currentLine++;
                    linesWritten++;
                    base.currentColumn--;
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
                    base.currentColumn++;
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[i][j].Symbol;
                    base.currentColumn++;
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[i][j].NominalValue;
                    base.currentColumn += 2;
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[i][j].TolerancePlus;
                    base.currentColumn++;
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[i][j].ToleranceMinus;
                    base.currentColumn++;
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[i][j].Value;

                    base.currentLine++;
                    linesWritten++;
                    base.currentColumn -= 6;

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
         * WriteHeader
         * 
         * Remplit l'entête du rapport Excel
         * 
         * header : Dictionary<string, string> - Dictionnaire contenant les informations de l'entête
         * designLine : int - Numéro de la ligne où écrire la désignation
         * 
         */
        public void WriteHeader(Dictionary<string, string> header, int designLine)
        {
            Excel.Worksheet ws = base.workbook.Sheets["Rapport d'essai dimensionnel"];

            ws.Cells[designLine, 4] = header["Designation"];
            ws.Cells[designLine + 2, 4] = header["N° de Plan"];
            ws.Cells[designLine + 4, 4] = header["Indice"];
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
    }
}
