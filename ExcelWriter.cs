using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Runtime.InteropServices;

enum LineType
{
    HEADER,
    MEASURE_TYPE,
    VALUE,
    VOID
}

namespace Application
{
    internal class ExcelWriter
    {
        private Excel.Application excelApp;
        private int currentLine;
        private char currentChar;

        [DllImport("Kernel32")]
        public static extern void AllocConsole();

        [DllImport("Kernel32", SetLastError = true)]
        public static extern void FreeConsole();

        public ExcelWriter() 
        {
            this.excelApp = new Excel.Application();
            this.excelApp.Workbooks.Open("C:\\Users\\LaboTri-PC2\\Desktop\\dev\\test\\test.xlsx");
            this.currentLine = 55;
            this.currentChar = 'A';

            AllocConsole();
        }

        ~ExcelWriter()
        {
            this.excelApp.Workbooks.Close();

            FreeConsole();
        }

        public void WriteData(List<String[]> data)
        {
            this.WriteExcelHeader();

            this.SkipHeader(data);

            for (int i = 0; i < data.Count; i++)
            {
                LineType type = this.GetLineType(data[i]);

                if(type == LineType.VALUE) 
                {
                    this.WriteValue(data[i]);
                }
            }
        }

        public void WriteExcelHeader()
        {
            excelApp.Range[currentChar + currentLine.ToString()].Value = "Nominal";

            this.currentChar += (char)2;

            excelApp.Range[currentChar + currentLine.ToString()].Value = "Tol.+";

            this.currentChar++;

            excelApp.Range[currentChar + currentLine.ToString()].Value = "Tol.-";

            currentChar = 'A';

            currentLine++;
        }

        public LineType GetLineType(String[] line)
        {
            if (line.Length == 0) return LineType.VOID;
            if (line[0] == "Designation") return LineType.HEADER;
            if (line[0][0] == '*') return LineType.MEASURE_TYPE;
            return LineType.VALUE;
        }

        public void WriteValue(String[] value)
        {
            for(int i = 0; i < value.Length; i++) 
            {
                Console.WriteLine(value[i]);
            }

            Console.WriteLine();
        }

        public void SkipHeader(List<String[]> data)
        {
            for(int i = 0; i < 6; i ++) 
            {
                data.RemoveAt(0);
            }
        }
    }
}
