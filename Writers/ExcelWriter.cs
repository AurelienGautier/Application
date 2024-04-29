using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Windows;
using Application.Data;
using Application.Exceptions;
using Microsoft.Office.Interop.Excel;

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

        private int rowToSign;
        private int colToSign;

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

            this.rowToSign = 51;
            this.colToSign = 14;

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

            this.setRowAndColFromFromType(_xlSheet);

            Clipboard.SetDataObject(image, true);
            var cellRngImg = (Excel.Range)_xlSheet.Cells[this.rowToSign, this.colToSign];
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

        private void setRowAndColFromFromType(Excel.Worksheet worksheet)
        {
            String? formType = worksheet.Cells[200, 1].Value;
            Console.WriteLine(formType);

            if (formType == null)
            {
                throw new ConfigDataException("Le type de formulaire n'a pas été reconnu.");
            }

            if (formType == "Rapport 1 pièce")
            {
                this.rowToSign = 55; this.colToSign = 14;
            }
            else if (formType == "Bague lisse" || formType == "Calibre à machoire" || formType == "Etalon colonne mesure" || formType == "Tampon lisse")
            {
                rowToSign = 52;
            }
        }

        /*-------------------------------------------------------------------------*/
    }
}