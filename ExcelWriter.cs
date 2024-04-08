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
        private char currentChar;
        private List<Piece> data;

        [DllImport("Kernel32")]
        public static extern void AllocConsole();

        [DllImport("Kernel32", SetLastError = true)]
        public static extern void FreeConsole();

        public ExcelWriter()
        {
            this.excelApp = new Excel.Application();
            this.workbook = this.excelApp.Workbooks.Open("C:\\Users\\LaboTri-PC2\\Desktop\\dev\\test\\test.xlsx");

            this.currentLine = 55;
            this.currentChar = 'A';
            this.data = new List<Piece>();

            AllocConsole();
        }

        ~ExcelWriter()
        {
            this.workbook.Close();

            this.excelApp.Quit();

            FreeConsole();
        }

        public void WriteData(List<Piece> data)
        {
            this.data = data;

            Console.WriteLine("oui");

            this.WriteExcelHeader();
            Console.WriteLine("non");

            this.WritePieceBaseValue();
            Console.WriteLine("peut-etre");

            this.WritePiecesValues();
            Console.WriteLine("padutou");
        }

        public void WriteExcelHeader()
        {
            WriteAndJump("N° Pièce", 0, (char)6);

            for (int i = 0; i < this.data.Count; i++)
            {
                this.WriteAndJump((i + 1).ToString(), 0, 1);
                this.WriteAndJump("Ecart", 0, 1);
                this.WriteAndJump("HT", 0, 1);
            }

            WriteAndJump("Observations", 1, -(6 + data.Count * 3));
            WriteAndJump("Nomnial", 0, 2);
            WriteAndJump("Tol.+", 0, 1);
            WriteAndJump("Tol.-", 0, 1);
            WriteAndJump("N° cote", 0, 1);
            WriteAndJump("N° M.C.", 1, -5);
        }

        public void WritePieceBaseValue()
        {
            this.data[0].WriteBaseValues(this.excelApp, this.currentChar, this.currentLine);
            this.currentChar += (char)6;
        }

        public void WritePiecesValues()
        {
            for(int i = 0; i < this.data.Count;i++)
            {
                this.data[i].WriteValues(excelApp, this.currentChar, this.currentLine);
                this.currentChar += (char)3;
            }
        }

        public void WriteAndJump(String thingToWrite, int lineJump, int columnJump)
        {
            this.excelApp.Range[this.currentChar + this.currentLine.ToString()].Value = thingToWrite;
            this.currentLine += lineJump;
            this.currentChar += (char)columnJump;
        }
    }
}