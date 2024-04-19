using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    internal class OnePieceWriter : ExcelWriter
    {
        private const int MAX_LINES = 22;

        public OnePieceWriter(string fileName, int firstLine, string formPath) : base(fileName, firstLine, 1, formPath)
        {
        }

        public override void CreateWorkSheets()
        {
            int linesToWrite = pieces[0].GetLinesToWriteNumber();

            int pageNumber = linesToWrite / MAX_LINES + 1;

            for (int i = 2; i <= pageNumber; i++)
            {
                workbook.Sheets["Mesures"].Copy(Type.Missing, workbook.Sheets[workbook.Sheets.Count]);
            }
        }

        public override void WritePiecesValues()
        {
            Excel.Worksheet ws = base.workbook.Sheets["Mesures"];

            List<String> measureTypes = pieces[0].GetMeasureTypes();
            List<List<Data.Data>> pieceData = pieces[0].GetData();

            int linesWritten = 0;
            int pageNumber = 1;

            for (int i = 0; i < pieceData.Count; i++)
            {
                // Écriture du plan
                if (measureTypes[i] != "")
                {
                    base.currentColumn++;
                    ws.Cells[base.currentLine, base.currentColumn].Value = measureTypes[i];
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

        public void WriteHeader(Dictionary<string, string> header, int designLine, int operatorLine)
        {
            Excel.Worksheet ws = base.workbook.Sheets["Rapport d'essai dimensionnel"];

            ws.Cells[designLine, 4] = header["Designation"];
            ws.Cells[designLine + 2, 4] = header["N° de Plan"];
            ws.Cells[designLine + 4, 4] = header["Indice"];
        }
    }
}
