﻿using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using Application.Data;
using Application.Exceptions;

namespace Application.Writers
{
    internal abstract class ExcelWriter
    {
        private readonly string fileToSavePath;

        protected Excel.Application excelApp;
        protected Excel.Workbook workbook;

        protected int currentLine;
        protected int currentColumn;
        protected List<Data.Piece> pieces;

        private int rowToSign;
        private int colToSign;

        protected bool modify;

        /*-------------------------------------------------------------------------*/

        /**
         * ExcelWriter
         * 
         * Constructeur de la classe
         * fileName : string - Chemin du fichier à sauvegarder
         * line : int - Ligne de la première cellule à remplir
         * col : int - Colonne de la première cellule à remplir
         * workBookPath : string - Chemin du formulaire vierge dans lequel écrire
         * 
         */
        protected ExcelWriter(string fileName, int line, int col, string workBookPath, bool modify)
        {
            this.fileToSavePath = fileName;
            this.excelApp = new Excel.Application();
            this.workbook = excelApp.Workbooks.Open(workBookPath);

            this.currentLine = line;
            this.currentColumn = col;

            this.rowToSign = 51;
            this.colToSign = 14;

            this.pieces = new List<Data.Piece>();
            this.modify = modify;
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
            var _xlSheet = (Excel.Worksheet)workbook.Sheets["Rapport d'essai dimensionnel"];

            this.setRowAndColFromType(_xlSheet);

            Clipboard.SetDataObject(ConfigSingleton.Instance.Signature, true);
            var cellRngImg = (Excel.Range)_xlSheet.Cells[this.rowToSign, this.colToSign];
            _xlSheet.Paste(cellRngImg, ConfigSingleton.Instance.Signature);
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
                this.fileToSavePath.Replace(".xlsx", ".pdf"),
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
                workbook.SaveAs(fileToSavePath);
            }
            catch
            {
                throw new Exceptions.ExcelFileAlreadyInUseException(this.fileToSavePath);
            }

            workbook.Close();
            excelApp.Quit();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * setRowAndColFromType
         * 
         * Détermine la ligne et la colonne où signer en fonction du type de formulaire
         * worksheet : Excel.Worksheet - Feuille de calculs du formulaire
         * 
         */
        private void setRowAndColFromType(Excel.Worksheet worksheet)
        {
            String? formType = worksheet.Cells[200, 1].Value;

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

        /**
         * EraseData
         * 
         * Efface les mesures présentes dans le fichier afin d'en écrire de nouvelles
         * 
         */
        public abstract void EraseData(int firstLine);

        /*-------------------------------------------------------------------------*/
    }
}