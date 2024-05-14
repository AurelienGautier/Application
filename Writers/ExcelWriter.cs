using Excel = Microsoft.Office.Interop.Excel;
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

        protected Form form;

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
        protected ExcelWriter(String fileName, Form form)
        {
            this.fileToSavePath = fileName;
            this.currentLine = form.FirstLine;
            this.currentColumn = form.FirstColumn;
            this.form = form;

            this.excelApp = new Excel.Application();
            this.workbook = excelApp.Workbooks.Open(form.Path);

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
        public void WriteData(List<Data.Piece> data)
        {
            this.pieces = data;

            writeHeader(data[0].GetHeader());

            CreateWorkSheets();

            WritePiecesValues();

            if (this.form.Sign)
            {
                signForm();

                exportFirstPageToPdf();
            }


            saveAndQuit();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * WriteHeader
         * 
         * Remplit l'entête du rapport Excel
         * 
         * header : Dictionary<string, string> - Dictionnaire contenant les informations de l'entête
         * designLine : int - Numéro de la ligne où écrire la désignation
         * 
         */
        private void writeHeader(Dictionary<string, string> header)
        {
            Excel.Worksheet ws = this.workbook.Sheets["Rapport d'essai dimensionnel"];

            ws.Cells[form.DesignLine, 4] = header["Designation"];
            ws.Cells[form.DesignLine + 2, 4] = header["N° de Plan"];
            ws.Cells[form.DesignLine + 4, 4] = header["Indice"];

            Excel.Workbook workbook2 = excelApp.Workbooks.Open(Environment.CurrentDirectory + "\\res\\ADRESSE");
            Excel.Worksheet ws2 = workbook2.Sheets["ADRESSE"];

            int currentLineWs2 = 2;

            // Tant que la ligne actuelle n'est pas vide et que le client n'a pas été trouvé
            while (ws2.Cells[currentLineWs2, 2].Value != null && ws2.Cells[currentLineWs2, 2].Value != header["Client"])
            {
                currentLineWs2++;
            }

            if (ws2.Cells[currentLineWs2, 2].Value == null)
            {
                return;
            }

            String address = ws2.Cells[currentLineWs2, 3].Value;
            String bp = ws2.Cells[currentLineWs2, 4].Value;
            String postalCode = ws2.Cells[currentLineWs2, 5].Value;
            String city = ws2.Cells[currentLineWs2, 6].Value;


            ws.Cells[form.ClientLine, 4] = header["Client"];
            ws.Cells[form.ClientLine + 1, 4] = address;
            ws.Cells[form.ClientLine + 2, 4] = bp;
            ws.Cells[form.ClientLine + 3, 4] = postalCode + " " + city;
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
        private void signForm()
        {
            var _xlSheet = (Excel.Worksheet)workbook.Sheets["Rapport d'essai dimensionnel"];

            Clipboard.SetDataObject(ConfigSingleton.Instance.Signature, true);
            var cellRngImg = (Excel.Range)_xlSheet.Cells[this.form.LineToSign, this.form.ColumnToSign];
            _xlSheet.Paste(cellRngImg, ConfigSingleton.Instance.Signature);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * ExportFirstPageToPdf
         * 
         * Exporte la première page du fichier excel en pdf (délégué aux classes filles)
         * 
         */
        private void exportFirstPageToPdf()
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
        private void saveAndQuit()
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
         * EraseData
         * 
         * Efface les mesures présentes dans le fichier afin d'en écrire de nouvelles
         * 
         */
        public abstract void EraseData(int firstLine);

        /*-------------------------------------------------------------------------*/
    }
}