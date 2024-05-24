using Application.Data;
using Application.Exceptions;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    /// <summary>
    /// Singleton class for managing the Excel API.
    /// </summary>
    internal class ExcelApiLinkSingleton
    {
        private static ExcelApiLinkSingleton? instance = null;
        private readonly Excel.Application excelApp;
        private readonly Dictionary<String, Excel.Workbook> workbooks;

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Returns the instance of the class, creating it if it is null.
        /// </summary>
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

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Constructor of the class.
        /// </summary>
        private ExcelApiLinkSingleton()
        {
            this.excelApp = new Excel.Application();
            this.workbooks = new Dictionary<String, Excel.Workbook>();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Destructor of the class.
        /// </summary>
        ~ExcelApiLinkSingleton()
        {
            foreach (var workbook in workbooks)
            {
                workbook.Value.Close();
            }

            excelApp.Quit();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Opens an Excel file and saves it in the list of open files.
        /// The file is identified by its path.
        /// </summary>
        /// <param name="path">Path of the file to open.</param>
        public void OpenWorkBook(String path)
        {
            if (!workbooks.ContainsKey(path))
            {
                try
                {
                    workbooks.Add(path, excelApp.Workbooks.Open(path));
                }
                catch
                {
                    throw new ConfigDataException("Le fichier " + path + " n'a pas été trouvé. Peut-être a-t-il été déplacé ou supprimé.");
                }
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Closes an open Excel file.
        /// The file is identified by its path.
        /// </summary>
        /// <param name="path">Path of the file to close.</param>
        public void CloseWorkBook(String path)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].Close();
                workbooks.Remove(path);
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Selects a worksheet in an open Excel file.
        /// </summary>
        /// <param name="path">Path of the file.</param>
        /// <param name="sheet">Number of the sheet to select.</param>
        public void ChangeWorkSheet(String path, int sheet)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].Sheets[sheet].Activate();
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Selects a worksheet in an open Excel file.
        /// </summary>
        /// <param name="path">Path of the file.</param>
        /// <param name="sheet">Name of the sheet to select.</param>
        public void ChangeWorkSheet(String path, String sheet)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].Sheets[sheet].Activate();
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Creates a new worksheet that is a copy of another worksheet.
        /// </summary>
        /// <param name="path">Path of the file.</param>
        /// <param name="sheetName">Name of the sheet to copy.</param>
        /// <param name="newSheetName">Name of the new sheet.</param>
        public void CopyWorkSheet(String path, String sheetName, String newSheetName)
        {
            if (!workbooks.ContainsKey(path)) return;


            try
            {
                workbooks[path].Sheets[workbooks[path].Sheets.Count].Name = newSheetName;
                workbooks[path].Sheets[sheetName].Copy(Type.Missing, workbooks[path].Sheets[workbooks[path].Sheets.Count]);
            }
            catch
            {
                // In case an exception is thrown if the sheet already exists, we simply want to do nothing
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes a value to a cell in a worksheet.
        /// </summary>
        /// <param name="path">Path of the file.</param>
        /// <param name="line">Line number.</param>
        /// <param name="column">Column number.</param>
        /// <param name="value">Value to write to the cell.</param>
        public void WriteCell(String path, int line, int column, String value)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].ActiveSheet.Cells[line, column] = value;
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Writes a value to a cell in a worksheet.
        /// </summary>
        /// <param name="path">Path of the file.</param>
        /// <param name="line">Line number.</param>
        /// <param name="column">Column number.</param>
        /// <param name="value">Value to write to the cell.</param>
        public void WriteCell(String path, int line, int column, double value)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].ActiveSheet.Cells[line, column] = value;
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Reads the value of a cell in a worksheet.
        /// </summary>
        /// <param name="path">Path of the file.</param>
        /// <param name="line">Line number.</param>
        /// <param name="column">Column number.</param>
        /// <returns>Value of the cell.</returns>
        public String ReadCell(String path, int line, int column)
        {
            if (workbooks.ContainsKey(path) && workbooks[path].ActiveSheet.Cells[line, column].Value != null)
            {
                return workbooks[path].ActiveSheet.Cells[line, column].Value.ToString();
            }

            return "";
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Merges cells in a worksheet.
        /// </summary>
        /// <param name="path">Path of the file.</param>
        /// <param name="line1">Line number of the first cell.</param>
        /// <param name="column1">Column number of the first cell.</param>
        /// <param name="line2">Line number of the second cell.</param>
        /// <param name="column2">Column number of the second cell.</param>
        public void MergeCells(String path, int line1, int column1, int line2, int column2)
        {
            if (!workbooks.ContainsKey(path)) return;
            
            Excel.Range range = workbooks[path].ActiveSheet.Range[
                workbooks[path].ActiveSheet.Cells[line1, column1],
                workbooks[path].ActiveSheet.Cells[line2, column2]];

            range.Merge();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Returns whether the specified range of cells is merged or not.
        /// </summary>
        /// <param name="path">Path of the file.</param>
        /// <param name="line1">Line number of the first cell.</param>
        /// <param name="column1">Column number of the first cell.</param>
        /// <param name="line2">Line number of the second cell.</param>
        /// <param name="column2">Column number of the second cell.</param>
        /// <returns>True if the cells are merged, false otherwise.</returns>
        public bool MergedCells(String path, int line1, int column1, int line2, int column2)
        {
            if(!workbooks.ContainsKey(path)) return false;

            Excel.Range range = workbooks[path].ActiveSheet.Range[
                workbooks[path].ActiveSheet.Cells[line1, column1],
                workbooks[path].ActiveSheet.Cells[line2, column2]];

            return range.MergeCells;
        }
        
        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Moves rows in a worksheet.
        /// </summary>
        /// <param name="path">Path of the file.</param>
        /// <param name="line">Line number to move.</param>
        /// <param name="startColumn">Number of the first column to move.</param>
        /// <param name="endColumn">Number of the last column to move.</param>
        /// <param name="linesToShift">Number of lines to move.</param>
        public void ShiftLines(String path, int line, int startColumn, int endColumn, int linesToShift)
        {
            if (!workbooks.ContainsKey(path)) return;

            for (int i = 0; i < linesToShift; i++)
            {
                workbooks[path].ActiveSheet.Range[
                    workbooks[path].ActiveSheet.Cells[line, startColumn],
                    workbooks[path].ActiveSheet.Cells[line, endColumn]]
                    .Insert(Excel.XlInsertShiftDirection.xlShiftDown, linesToShift);
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Returns the address of a cell.
        /// </summary>
        /// <param name="row">Row number.</param>
        /// <param name="col">Column number.</param>
        /// <returns>Address of the cell.</returns>
        public String GetCellAddress(int row, int col)
        {
            if (col <= 0 || row <= 0)
            {
                throw new ArgumentException("what you pass as parameter is shit");
            }

            int dividend = col;
            StringBuilder columnName = new StringBuilder();

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName.Insert(0, Convert.ToChar('A' + modulo));
                dividend = (dividend - modulo) / 26;
            }

            return columnName.ToString() + row.ToString();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Pastes an image into a cell.
        /// </summary>
        /// <param name="path">Path of the Excel file.</param>
        /// <param name="line">Line number to place the image.</param>
        /// <param name="column">Column number to place the image.</param>
        /// <param name="image">Image to paste.</param>
        public void PasteImage(String path, int line, int column, Image image)
        {
            if (!workbooks.ContainsKey(path)) return;

            Clipboard.SetDataObject(image, true);
            var cellRngImg = (Excel.Range)this.workbooks[path].ActiveSheet.Cells[line, column];
            this.workbooks[path].ActiveSheet.Paste(cellRngImg, ConfigSingleton.Instance.Signature);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Exports the first page of an Excel file to PDF.
        /// </summary>
        /// <param name="path">Path of the Excel file.</param>
        /// <param name="pdfPath">Path of the PDF file to export.</param>
        public void ExportFirstPageToPdf(String path, String pdfPath)
        {
            if (!workbooks.ContainsKey(path)) return;

            workbooks[path].ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfPath, Type.Missing, Type.Missing, Type.Missing, 1, 1, false, Type.Missing);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Saves an Excel file.
        /// </summary>
        /// <param name="path">Path of the file to save.</param>
        /// <param name="pathToSave">Path to save the file.</param>
        public void SaveWorkBook(String path, String pathToSave)
        {
            if (!workbooks.ContainsKey(path)) return;

            this.workbooks[path].Sheets[1].Activate();
            try
            {
                workbooks[path].SaveAs(pathToSave);
            }
            catch
            {
                throw new Exceptions.ExcelFileAlreadyInUseException(pathToSave);
            }
        }

        /*-------------------------------------------------------------------------*/
    }
}
