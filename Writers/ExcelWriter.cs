using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Windows;
using Application.Data;

namespace Application.Writers
{
    internal abstract class ExcelWriter
    {
        private readonly string fileToSaveName;

        protected Excel.Application excelApp;
        protected Excel.Workbook workbook;

        protected int currentLine;
        protected int currentColumn;
        protected List<Data.Piece> pieces;

        /*-------------------------------------------------------------------------*/

        /**
         * ExcelWriter
         * 
         * Constructeur de la classe
         * fileName : string - Nom du fichier à sauvegarder
         * line : int - Ligne de la première cellule à remplir
         * col : int - Colonne de la première cellule à remplir
         * workBookPath : string - Chemin du formulaire vierge dans lequel écrire
         * 
         */
        protected ExcelWriter(string fileName, int line, int col, string workBookPath)
        {
            this.fileToSaveName = fileName;
            this.excelApp = new Excel.Application();
            this.workbook = excelApp.Workbooks.Open(workBookPath);

            this.currentLine = line;
            this.currentColumn = col;

            this.pieces = new List<Data.Piece>();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * WriteData
         * 
         * Ecrit les données des pièces dans le fichier excel
         * data : List<Piece> - Liste des pièces à écrire
         * 
         */
        public void WriteData(List<Data.Piece> data, bool sign)
        {
            this.pieces = data;

            CreateWorkSheets();

            WritePiecesValues();

            if (sign)
            {
                SignForm();

                ExportFirstPageToPdf();
            }


            SaveAndQuit();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * CreateWorkSheets
         * 
         * Crée les feuilles de calculs nécessaires (délégué aux classes filles)
         * 
         */
        public abstract void CreateWorkSheets();

        /*-------------------------------------------------------------------------*/

        /**
         * WritePiecesValues
         * 
         * Ecrit les valeurs des pièces dans le fichier excel (délégué aux classes filles)
         * 
         */
        public abstract void WritePiecesValues();

        /*-------------------------------------------------------------------------*/

        /**
         * SignForm
         * 
         * Signe le formulaire
         * 
         */
        public void SignForm()
        {
            Image image;

            try
            {
                image = Image.FromFile(ConfigSingleton.Instance.Signature);
            }
            catch
            {
                throw new System.ArgumentException("Chemin vers la signature vide ou incorrect");
            }

            var _xlSheet = (Excel.Worksheet)workbook.Sheets["Rapport d'essai dimensionnel"];

            Clipboard.SetDataObject(image, true);
            var cellRngImg = (Excel.Range)_xlSheet.Cells[55, 14];
            _xlSheet.Paste(cellRngImg, image);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * ExportFirstPageToPdf
         * 
         * Exporte la première page du fichier excel en pdf (délégué aux classes filles)
         * 
         */
        public void ExportFirstPageToPdf()
        {
            this.workbook.ExportAsFixedFormat(
                Excel.XlFixedFormatType.xlTypePDF, 
                this.fileToSaveName.Replace(".xlsx", ".pdf"),
                Type.Missing,
                Type.Missing,
                Type.Missing,
                1,
                1,
                false,
                Type.Missing
            );
        }

        /*-------------------------------------------------------------------------*/

        /**
         * SaveAndQuit
         * 
         * Sauvegarde le fichier et ferme l'application
         * 
         */
        public void SaveAndQuit()
        {
            this.workbook.Sheets[1].Activate();

            try
            {
                workbook.SaveAs(fileToSaveName);
            }
            catch
            {
                throw new Exceptions.ExcelFileAlreadyInUseException(this.fileToSaveName);
            }

            workbook.Close();
            excelApp.Quit();
        }

        /*-------------------------------------------------------------------------*/
    }
}