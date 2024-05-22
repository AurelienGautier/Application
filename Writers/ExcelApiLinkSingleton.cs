using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    internal class ExcelApiLinkSingleton
    {
        private static ExcelApiLinkSingleton? instance = null;
        Excel.Application excelApp;
        Dictionary<String, Excel.Workbook> workbooks;
        public static ExcelApiLinkSingleton Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new ExcelApiLinkSingleton();
                }

                return instance;
            }
        }

        private ExcelApiLinkSingleton()
        {
            this.excelApp = new Excel.Application();
            this.workbooks = new Dictionary<String, Excel.Workbook>();
        }

        ~ExcelApiLinkSingleton()
        {
            foreach (var workbook in workbooks)
            {
                workbook.Value.Close();
            }

            excelApp.Quit();
        }

        public void OpenWorkBook(String path)
        {
            if (!workbooks.ContainsKey(path))
            {
                workbooks.Add(path, excelApp.Workbooks.Open(path));
            }
        }

        public void CloseWorkBook(String path)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].Close();
                workbooks.Remove(path);
            }
        }

        public void ChangeWorkSheet(String path, int sheet)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].Sheets[sheet].Activate();
            }
        }

        public void WriteCell(String path, int line, int column, String value)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].ActiveSheet.Cells[line, column] = value;
            }
        }

        public String? ReadCell(String path, int line, int column)
        {
            if (workbooks.ContainsKey(path) && workbooks[path].ActiveSheet.Cells[line, column].Value != null)
            {
                return workbooks[path].ActiveSheet.Cells[line, column].Value.ToString();
            }

            return null;
        }

        public void MergeCells(String path, int line1, int column1, int line2, int column2)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].ActiveSheet.Range[
                    workbooks[path].ActiveSheet.Cells[line1, column1],
                    workbooks[path].ActiveSheet.Cells[line2, column2]]
                    .Merge();
            }
        }

        public void ShiftLines(String path, int line, int startColumn, int endColumn, int linesToShift)
        {
            if (!workbooks.ContainsKey(path)) return;

            for(int i = 0; i < linesToShift; i++)
            {
                workbooks[path].ActiveSheet.Range[
                    workbooks[path].ActiveSheet.Cells[line, startColumn],
                    workbooks[path].ActiveSheet.Cells[line, endColumn]]
                    .Insert(Excel.XlInsertShiftDirection.xlShiftDown, linesToShift);
            }
        }

        public String GetCellAddress(int row, int col)
        {
            if (col <= 0 || row <= 0)
            {
                throw new ArgumentException("quoi toi passer en paramètre être merde");
            }

            int dividend = col;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName + row.ToString();
        }
    }
}
