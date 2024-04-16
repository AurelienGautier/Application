﻿using System;
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
            int pageNumber = pieces[0].GetLinesToWriteNumber() / MAX_LINES + 1;

            int iterations = base.pieces.Count / 5;
            if (base.pieces.Count % 5 != 0) iterations++;

            pageNumber *= iterations;

            for(int i = 4; i <= pageNumber; i++)
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

            int iterations = pieceData.Count / 5;
            if(pieceData.Count % 5 != 0) iterations++;

            for (int i = 0; i < iterations; i++)
            {
                this.Write5pieces();

                this.min += 5;

                if (i == pieceData.Count / 5 - 1 && pieceData.Count % 5 != 0) this.max = pieceData.Count;
                else this.max += 5;

                if(i < iterations - 1) this.ChangePage();
            }
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
                if (this.linesWritten == MAX_LINES) { this.ChangePage(); }

                for (int j = 0; j < pieceData[0][i].Count; j++)
                {
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[0][i][j].GetNominalValue();
                    ws.Cells[base.currentLine, base.currentColumn + 2].Value = pieceData[0][i][j].GetTolPlus();
                    ws.Cells[base.currentLine, base.currentColumn + 3].Value = pieceData[0][i][j].GetTolMinus();

                    base.currentColumn += 3;

                    // Écriture des valeurs des pièces
                    for(int k = this.min; k < this.max; k++)
                    {
                        base.currentColumn += 3;
                        ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[k][i][j].GetValue();
                    }

                    base.currentColumn -= (3 + 3 * (this.max - this.min));

                    base.currentLine++;
                    this.linesWritten++;

                    // Changement de page si l'actuelle est complète ou si arrivé à la fin des 5 pièces
                    if (this.linesWritten == MAX_LINES) { this.ChangePage(); }
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

            for (int i = this.min; i < this.min + 5; i++)
            {
                ws.Cells[15, col].Value = i + 1;
                col += 3;
            }
        }
    }
}
