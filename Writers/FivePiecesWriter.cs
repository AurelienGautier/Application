using System;
using System.Collections.Generic;
using System.Linq;
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
