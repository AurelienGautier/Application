﻿using Excel = Microsoft.Office.Interop.Excel;

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

            this.pieces = new List<Piece>();
        }

        /**
         * WriteData
         * 
         * Ecrit les données des pièces dans le fichier excel
         * data : List<Piece> - Liste des pièces à écrire
         * 
         */
        public void WriteData(List<Piece> data)
        {
            this.pieces = data;

            CreateWorkSheets();

            WritePiecesValues();

            SaveAndQuit();
        }

        /**
         * CreateWorkSheets
         * 
         * Crée les feuilles de calculs nécessaires (délégué aux classes filles)
         * 
         */
        public abstract void CreateWorkSheets();

        /**
         * WritePiecesValues
         * 
         * Ecrit les valeurs des pièces dans le fichier excel (délégué aux classes filles)
         * 
         */
        public abstract void WritePiecesValues();

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
                throw new Exceptions.ExcelFileAlreadyInUseException();
            }

            workbook.Close();
            excelApp.Quit();
        }
    }
}