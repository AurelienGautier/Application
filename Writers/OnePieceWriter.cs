using Microsoft.Office.Interop.Excel;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    internal class OnePieceWriter : ExcelWriter
    {
        private const int MAX_LINES = 22;

        public OnePieceWriter(string fileName, int firstLine, string formPath) : base(fileName, firstLine, 1, formPath)
        {
        }

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
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[i][j].GetSymbol();
                    base.currentColumn++;
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[i][j].GetNominalValue();
                    base.currentColumn += 2;
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[i][j].GetTolPlus();
                    base.currentColumn++;
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[i][j].GetTolMinus();
                    base.currentColumn++;
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[i][j].GetValue();

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

        /**
         * ExcportFirstPageToPdf
         * 
         * Exporte la première page du formulaire Excel en PDF
         * 
         */
        public override void ExportFirstPageToPdf()
        {
        }
    }
}
