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
        Excel.Worksheet ws;
        int min;
        int max;
        const int MAX_LINES = 23;

        public FivePiecesWriter(string fileName) : base(fileName, 17, 1, "C:\\Users\\LaboTri-PC2\\Desktop\\dev\\form\\rapport5pieces")
        {
            this.pageNumber = 1;
            this.measureTypes = new List<List<String>>();
            this.pieceData = new List<List<List<Data.Data>>>();
            this.linesWritten = 0;
            this.ws = base.workbook.Sheets["Mesures"];
            this.min = 0;
            this.max = 5;
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
            for(int i = 0; i < base.pieces.Count; i++)
            {
                this.measureTypes.Add(base.pieces[i].GetMeasureTypes());
                this.pieceData.Add(base.pieces[i].GetData());
            }

            for (int i = 0; i < pieceData.Count / 5; i++)
            {
                this.Write5pieces();

                this.min += 5;
                this.max += 5;
                if(i < pieceData.Count / 5 - 1) this.ChangePage();
            }

            this.ws = base.workbook.Sheets["Mesures"];
        }

        public void Write5pieces()
        {
            for (int i = 0; i < pieceData[0].Count; i++)
            {
                // Écriture du plan
                if (measureTypes[0][i] != "")
                {
                    ws.Cells[base.currentLine, base.currentColumn].Value = measureTypes[0][i];
                    base.currentLine++;
                    this.linesWritten++;
                }

                // Changement de page si l'actuelle est complète
                if (this.linesWritten == MAX_LINES) this.ChangePage();

                for (int j = 0; j < pieceData[0][i].Count; j++)
                {
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[0][i][j].GetNominalValue();
                    ws.Cells[base.currentLine, base.currentColumn + 2].Value = pieceData[0][i][j].GetTolPlus();
                    ws.Cells[base.currentLine, base.currentColumn + 3].Value = pieceData[0][i][j].GetTolMinus();

                    base.currentLine++;
                    this.linesWritten++;

                    // Écriture des valeurs des pièces
                    for(int k = this.min; k < this.max; k++)
                    {

                    }

                    // Changement de page si l'actuelle est complète ou si arrivé à la fin des 5 pièces
                    if (this.linesWritten == MAX_LINES) this.ChangePage();
                }
            }
        }

        public void ChangePage()
        {
            this.pageNumber++;

            ws = this.workbook.Sheets["Mesures (" + this.pageNumber.ToString() + ")"];

            base.currentLine = 17;
            this.linesWritten = 0;

            int col = 7;

            for (int i = this.min; i < this.max; i++)
            {
                ws.Cells[15, col].Value = i + 1;
                col += 3;
            }
        }

        public int GetWorksheetNumberToCreate()
        {
            int lineNumber = 0;

            int min = 0;
            int max = 5;

            while (max <= base.pieces.Count)
            {
                int temp = 0;
                for (int i = min; i < max; i++)
                {
                    temp += pieces[i].GetLinesToWriteNumber();
                }

                temp = temp / MAX_LINES + 1;
                temp = temp / 5;

                lineNumber += temp;

                min += 5;
                max += 5;
            }

            return lineNumber;
        }
    }
}
