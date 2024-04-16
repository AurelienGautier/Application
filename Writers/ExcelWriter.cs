using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.IO;

namespace Application.Writers
{
    internal abstract class ExcelWriter
    {
        private readonly string fileToSaveName;

        protected Excel.Application excelApp;
        protected Excel.Workbook workbook;

        protected int currentLine;
        protected int currentColumn;
        protected List<Piece> pieces;

        protected ExcelWriter(string fileName, int line, int col, string workBookPath)
        {
            fileToSaveName = fileName;
            excelApp = new Excel.Application();
            workbook = excelApp.Workbooks.Open(workBookPath);

            currentLine = line;
            currentColumn = col;

            pieces = new List<Piece>();
        }

        public void WriteData(List<Piece> data)
        {
            pieces = data;

            CreateWorkSheets();

            WritePiecesValues();

            SaveAndQuit();
        }

        public abstract void CreateWorkSheets();

        public abstract void WritePiecesValues();

        public void SaveAndQuit()
        {
            this.workbook.Sheets[1].Activate();

            try
            {
                workbook.SaveAs(fileToSaveName);
            }
            catch
            {
                throw new Exceptions.ExcelFileAlreadyInUseException();
            }

            workbook.Close();
            excelApp.Quit();
        }
    }
}