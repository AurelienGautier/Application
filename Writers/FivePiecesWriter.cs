using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    internal class FivePiecesWriter : ExcelWriter
    {
        public FivePiecesWriter(string fileName) : base(fileName, 17, 1, "C:\\Users\\LaboTri-PC2\\Desktop\\dev\\form\\rapport5pieces")
        {
        }

        public override void CreateWorkSheets()
        {
            int workSheetNumber = GetWorksheetNumberToCreate();

            for(int i = 0; i < workSheetNumber; i++)
            {
                workbook.Sheets["Mesures"].Copy(Type.Missing, workbook.Sheets[workbook.Sheets.Count]);
            }
        }

        public override void WritePiecesValues()
        {
            Excel.Worksheet ws = base.workbook.Sheets["Mesures"];

            List<List<String>> measureTypes = new List<List<String>>();
            List<List<List<Data.Data>>> pieceData = new List<List<List<Data.Data>>>();

            for(int i = 0; i < base.pieces.Count; i++)
            {
                measureTypes.Add(base.pieces[i].GetMeasureTypes());
                pieceData.Add(base.pieces[i].GetData());
            }

            int linesWritten = 0;
            int pageNumber = 1;

            for(int i = 0; i < pieceData[0].Count; i++)
            {
                // Écriture du plan
                if (measureTypes[0][i] != "")
                {
                    ws.Cells[base.currentLine, base.currentColumn].Value = measureTypes[0][i];
                    base.currentLine++;
                    linesWritten++;
                }

                // Changement de page si l'actuelle est complète
                if (linesWritten == 22)
                {
                    pageNumber++;

                    ws = this.workbook.Sheets["Mesures (" + pageNumber.ToString() + ")"];

                    base.currentLine -= linesWritten;
                    linesWritten = 0;
                }

                for (int j = 0; j < pieceData[0][i].Count; j++)
                {
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[0][i][j].GetNominalValue();
                    base.currentColumn+=2;
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[0][i][j].GetTolPlus();
                    base.currentColumn++;
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[0][i][j].GetTolMinus();
                    base.currentLine++;
                    linesWritten++;

                    // à enlever après
                    base.currentColumn -= 3;

                    // Changement de page si l'actuelle est complète
                    if (linesWritten == 22 || j == pieceData[0][i].Count)
                    {
                        pageNumber++;

                        ws = this.workbook.Sheets["Mesures (" + pageNumber.ToString() + ")"];

                        base.currentLine -= linesWritten;
                        linesWritten = 0;
                    }
                }
            }
        }

        public int GetWorksheetNumberToCreate()
        {
            int lineNumber = 0;

            int min = 0;
            int max = 5;

            while(max <= base.pieces.Count)
            {
                int temp = 0;
                for (int i = min; i < max; i++)
                {
                    temp += pieces[i].GetLinesToWriteNumber();
                }

                temp = temp / 22 + 1;
                temp = temp / 5;

                lineNumber += temp;

                min += 5;
                max += 5;
            }

            return lineNumber;
        }
    }
}
