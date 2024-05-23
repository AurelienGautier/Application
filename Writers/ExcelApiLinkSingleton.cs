using Application.Data;
using System.Drawing;
using System.Text;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    internal class ExcelApiLinkSingleton
    {
        private static ExcelApiLinkSingleton? instance = null;
        private readonly Excel.Application excelApp;
        private readonly Dictionary<String, Excel.Workbook> workbooks;

        /*-------------------------------------------------------------------------*/

        /**
         * Retourne l'instance de la classe en la créant si elle est à null
         * 
         */
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

        /**
         * Constructeur de la classe         
         * 
         */
        private ExcelApiLinkSingleton()
        {
            this.excelApp = new Excel.Application();
            this.workbooks = new Dictionary<String, Excel.Workbook>();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Destructeur de la classe
         * 
         * Libère tous les workbooks ouverts et ferme l'application Excel
         * 
         */
        ~ExcelApiLinkSingleton()
        {
            foreach (var workbook in workbooks)
            {
                workbook.Value.Close();
            }

            excelApp.Quit();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Permet d'ouvrir un fichier excel et de le sauvegarder dans la liste des fichiers ouverts
         * Le fichier est identifiable par son chemin
         * 
         * path : String - Chemin du fichier à ouvrir
         * 
         */
        public void OpenWorkBook(String path)
        {
            if (!workbooks.ContainsKey(path))
            {
                workbooks.Add(path, excelApp.Workbooks.Open(path));
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Permet de fermer un fichier excel ouvert
         * Le fichier est identifiable par son chemin
         * 
         * path : String - Chemin du fichier à fermer
         * 
         */
        public void CloseWorkBook(String path)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].Close();
                workbooks.Remove(path);
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Sélectionne une feuille de calcul dans un fichier excel ouvert
         * 
         * path : String - Chemin du fichier
         * sheet : int - Numéro de la feuille à sélectionner
         * 
         */
        public void ChangeWorkSheet(String path, int sheet)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].Sheets[sheet].Activate();
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Sélectionne une feuille de calcul dans un fichier excel ouvert
         * 
         * path : String - Chemin du fichier
         * sheet : String - Nom de la feuille à sélectionner
         * 
         */
        public void ChangeWorkSheet(String path, String sheet)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].Sheets[sheet].Activate();
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Créer une nouvelle feuille de calcul qui est une copie d'une autre
         * 
         * path : String - Chemin du fichier
         * sheetName : String - Nom de la feuille à copier
         * newSheetName : String - Nom de la nouvelle feuille
         * 
         */
        public void CopyWorkSheet(String path, String sheetName, String newSheetName)
        {
            if (!workbooks.ContainsKey(path)) return;
            
            workbooks[path].Sheets[sheetName].Copy(Type.Missing, workbooks[path].Sheets[workbooks[path].Sheets.Count]);

            try
            {
                workbooks[path].Sheets[workbooks[path].Sheets.Count].Name = newSheetName;
            }
            catch
            {
                // Dans le cas où une exception est levée si la feuille existe déjà, on souhaite simplement ne rien faire
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Écrit une valeur dans une cellule d'une feuille de calcul
         * 
         * path : String - Chemin du fichier
         * line : int - Numéro de la ligne
         * column : int - Numéro de la colonne
         * value : String - Valeur à écrire dans la cellule
         * 
         */
        public void WriteCell(String path, int line, int column, String value)
        {
            if (workbooks.ContainsKey(path))
            {
                workbooks[path].ActiveSheet.Cells[line, column] = value;
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Lit la valeur d'une cellule d'une feuille de calcul
         * 
         * path : String - Chemin du fichier
         * line : int - Numéro de la ligne
         * column : int - Numéro de la colonne
         * 
         * return : String - Valeur de la cellule
         * 
         */
        public String ReadCell(String path, int line, int column)
        {
            if (workbooks.ContainsKey(path) && workbooks[path].ActiveSheet.Cells[line, column].Value != null)
            {
                return workbooks[path].ActiveSheet.Cells[line, column].Value.ToString();
            }

            return "";
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Fusionne des cellules d'une feuille de calcul
         * 
         * path : String - Chemin du fichier
         * line1 : int - Numéro de ligne de la première cellule
         * column1 : int - Numéro de colonne de la première cellule
         * line2 : int - Numéro de ligne de la deuxième cellule
         * column2 : int - Numéro de colonne de la deuxième cellule
         * 
         */
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

        /*-------------------------------------------------------------------------*/

        /**
         * Déplace des lignes d'une feuille de calcul
         * 
         * path : String - Chemin du fichier
         * line : int - Numéro de la ligne à déplacer
         * startColumn : int - Numéro de la première colonne à déplacer
         * endColumn : int - Numéro de la dernière colonne à déplacer
         * linesToShift : int - Nombre de lignes à déplacer
         * 
         */
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

        /**
         * Retourne l'adresse d'une cellule
         * 
         * row : int - Numéro de la ligne
         * col : int - Numéro de la colonne
         * 
         * return : String - Adresse de la cellule
         * 
         */
        public String GetCellAddress(int row, int col)
        {
            if (col <= 0 || row <= 0)
            {
                throw new ArgumentException("quoi toi passer en paramètre être merde");
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

        /**
         * Colle une image dans une cellule
         * 
         * path : String - Chemin du fichier excel
         * line : int - Numéro de la ligne où mettre l'image
         * column : int - Numéro de la colonne où mettre l'image
         * image : Image - Image à coller
         * 
         */
        public void PasteImage(String path, int line, int column, Image image)
        {
            if (!workbooks.ContainsKey(path)) return;

            Clipboard.SetDataObject(image, true);
            var cellRngImg = (Excel.Range)this.workbooks[path].ActiveSheet.Cells[line, column];
            this.workbooks[path].ActiveSheet.Paste(cellRngImg, ConfigSingleton.Instance.Signature);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Exporte la première page d'un fichier excel en pdf
         * 
         * path : String - Chemin du fichier excel
         * pdfPath : String - Chemin du fichier pdf à exporter
         * 
         */
        public void ExportFirstPageToPdf(String path, String pdfPath)
        {
            if (!workbooks.ContainsKey(path)) return;

            workbooks[path].ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfPath, Type.Missing, Type.Missing, Type.Missing, 1, 1, false, Type.Missing);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Sauvegarde un fichier excel
         * 
         * path : String - Chemin du fichier à sauvegarder
         * pathToSave : String - Chemin où sauvegarder le fichier
         * 
         */
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
