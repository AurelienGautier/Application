using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    internal class JawGauge : ExcelWriter
    {
        public JawGauge(string fileName) : base(fileName, 17, 1, "C:\\Users\\LaboTri-PC2\\Desktop\\dev\\form\\calibreAmachoire")
        {
        }

        public override void CreateWorkSheets()
        {
        }

        public override void WritePiecesValues()
        {
        }

        public override void ExportFirstPageToPdf()
        {
        }
    }
}
