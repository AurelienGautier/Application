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
        int pageNumber;
        List<List<String>> measureTypes;
        List<List<List<Data.Data>>> pieceData;
        int linesWritten;

        public FivePiecesWriter(string fileName) : base(fileName, 17, 1, "C:\\Users\\LaboTri-PC2\\Desktop\\dev\\form\\rapport5pieces")
        {
            this.pageNumber = 1;
            this.measureTypes = new List<List<String>>();
            this.pieceData = new List<List<List<Data.Data>>>();
            this.linesWritten = 0;
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

            for(int i = 0; i < base.pieces.Count; i++)
            {
                this.measureTypes.Add(base.pieces[i].GetMeasureTypes());
                this.pieceData.Add(base.pieces[i].GetData());
            }

            for(int k = 0; k < pieceData.Count / 5; k++)
            {
                this.Write5pieces(ws);
            }

        }

        public void Write5pieces(Excel.Worksheet ws)
        {
            int linesWritten = 0;

            for (int i = 0; i < pieceData[0].Count; i++)
            {
                // Écriture du plan
                if (measureTypes[0][i] != "")
                {
                    ws.Cells[base.currentLine, base.currentColumn].Value = measureTypes[0][i];
                    base.currentLine++;
                    linesWritten++;
                }

                // Changement de page si l'actuelle est complète
                if (linesWritten == 22) this.ChangePage(ws);

                for (int j = 0; j < pieceData[0][i].Count; j++)
                {
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[0][i][j].GetNominalValue();
                    ws.Cells[base.currentLine, base.currentColumn + 2].Value = pieceData[0][i][j].GetTolPlus();
                    ws.Cells[base.currentLine, base.currentColumn + 3].Value = pieceData[0][i][j].GetTolMinus();
                    base.currentLine++;
                    this.linesWritten++;

                    // Changement de page si l'actuelle est complète ou si arrivé à la fin des 5 pièces
                    if (this.linesWritten == 22 || j == pieceData[0][i].Count - 1) this.ChangePage(ws);
                }
            }
        }

        public void ChangePage(Excel.Worksheet ws)
        {
            this.pageNumber++;

            Console.WriteLine("Page number: " + this.pageNumber);

            ws = this.workbook.Sheets["Mesures (" + this.pageNumber.ToString() + ")"];

            base.currentLine -= linesWritten;
            this.linesWritten = 0;
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
