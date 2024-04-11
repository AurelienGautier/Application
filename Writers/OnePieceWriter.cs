using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    internal class OnePieceWriter : ExcelWriter
    {
        public OnePieceWriter(string fileName) : base(fileName, 30, 1, "C:\\Users\\LaboTri-PC2\\Desktop\\dev\\form\\rapport1piece")
        {
        }

        public override void CreateWorkSheets()
        {
            int linesToWrite = pieces[0].GetLinesToWriteNumber();

            int pageNumber = linesToWrite / 22 + 1;

            Excel.Worksheet ws = workbook.Sheets["Mesures"];

            for (int i = 4; i <= pageNumber; i++)
            {
                workbook.Sheets["Mesures"].Copy(Type.Missing, workbook.Sheets[workbook.Sheets.Count]);
            }
        }

        public override void WritePiecesValues()
        {
            for (int i = 0; i < pieces.Count; i++)
            {
                pieces[i].WriteValues(workbook, currentLine, currentColumn);
            }
        }
    }
}
