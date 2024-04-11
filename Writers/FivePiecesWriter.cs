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
            
        }

        public override void WritePiecesValues()
        {
            
        }

        public int GetLineNumberToCreate()
        {
            int lineNumber = 0;

            lineNumber = base.pieces.Count / 5 + 1;

            return lineNumber;
        }
    }
}
