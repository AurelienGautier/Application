using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Application
{
    internal class ExcelWriter
    {
        private Excel.Application excelApp;
        private Excel.Workbook workbook;

        private int currentLine;
        private int currentColumn;
        private List<Piece> pieces;

        public ExcelWriter()
        {
            this.excelApp = new Excel.Application();
            this.workbook = this.excelApp.Workbooks.Open("C:\\Users\\LaboTri-PC2\\Desktop\\dev\\form\\rapport1piece");

            this.currentLine = 30;
            this.currentColumn = 1;

            this.pieces = new List<Piece>();
        }

        ~ExcelWriter()
        {
            
        }

        public void WriteData(List<Piece> data)
        {
            this.pieces = data;

            this.CreateWorkSheets();

            this.WritePiecesValues();

            this.workbook.SaveAs("C:\\Users\\LaboTri-PC2\\Desktop\\dev\\test\\rappport1piece");

            this.workbook.Close();

            this.excelApp.Quit();
        }

        public void CreateWorkSheets()
        {
            int linesToWrite = this.pieces[0].GetLinesToWriteNumber();

            int pageNumber = linesToWrite / 22 + 1;

            Excel.Worksheet ws = this.workbook.Sheets["Mesures"];

            for (int i = 4; i <= pageNumber; i++)
            {
                this.workbook.Sheets["Mesures"].Copy(Type.Missing, this.workbook.Sheets[this.workbook.Sheets.Count]);
            }
        }

        public void WritePieceBaseValue()
        {
            this.pieces[0].WriteBaseValues(this.workbook, this.currentLine, this.currentColumn);
            this.currentColumn+= 6;
        }

        public void WritePiecesValues()
        {
            for(int i = 0; i < this.pieces.Count;i++)
            {
                this.pieces[i].WriteValues(this.workbook, this.currentLine, this.currentColumn);
                this.currentColumn += 3;
            }
        }
    }
}