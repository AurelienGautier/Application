using Application.Data;
using Application.Exceptions;

namespace Application.Writers
{
    internal abstract class ExcelWriter
    {
        private readonly string fileToSavePath;
        protected int currentLine;
        protected int currentColumn;
        protected List<Data.Piece> pieces;
        protected Form form;
        protected ExcelApiLinkSingleton excelApiLink;

        /*-------------------------------------------------------------------------*/

        /**
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

            ExcelApiLinkSingleton.Instance.OpenWorkBook(form.Path);

            this.pieces = new List<Data.Piece>();

            this.excelApiLink = ExcelApiLinkSingleton.Instance;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Ecrit les données des pièces dans le fichier excel
         * data : List<Piece> - Liste des pièces à écrire
         * 
         */
        public void WriteData(List<Data.Piece> data, List<Standard> standards)
        {
            this.pieces = data;

            writeHeader(data[0].GetHeader(), standards);

            if (!form.Modify) CreateWorkSheets();

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
         * Remplit l'en-tête du rapport Excel
         * 
         * header : Dictionary<string, string> - Dictionnaire contenant les informations de l'entête
         * designLine : int - Numéro de la ligne où écrire la désignation
         * 
         */
        private void writeHeader(Header header, List<Standard> standards)
        {
            excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["HeaderPage"]);

            excelApiLink.WriteCell(form.Path, form.DesignLine, 4, header.Designation);
            excelApiLink.WriteCell(form.Path, form.DesignLine + 2, 4, header.PlanNb);
            excelApiLink.WriteCell(form.Path, form.DesignLine + 4, 4, header.Index);
            excelApiLink.WriteCell(form.Path, 14, 1, "N° " + header.ObservationNum);
            excelApiLink.WriteCell(form.Path, 38, 8, header.PieceReceptionDate);
            excelApiLink.WriteCell(form.Path, 40, 4, header.Observations);

            this.writeClient(header.ClientName);

            if(!form.Modify) this.writeStandards(standards);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Remplit les informations du client dans l'en-tête du formulaire en allant les chercher dans le fichier des clients
         * 
         * client : String - Nom du client
         * 
         */
        private void writeClient(String client)
        {
            String clientWorkbookPath = Environment.CurrentDirectory + "\\res\\ADRESSE";
            excelApiLink.OpenWorkBook(clientWorkbookPath);
            excelApiLink.ChangeWorkSheet(clientWorkbookPath, "ADRESSE");

            int currentLineWs2 = 2;

            // Tant que la ligne actuelle n'est pas vide et que le client n'a pas été trouvé
            while (excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 2) != ""
                && excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 2) != client)
            {
                currentLineWs2++;
            }

            String address = "";
            String bp = "";
            String postalCode = "";
            String city = "";

            if (excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 2) != "")
            {
                address = excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 3);
                bp = excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 4);
                postalCode = excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 5);
                city = excelApiLink.ReadCell(clientWorkbookPath, currentLineWs2, 6);
            }

            excelApiLink.CloseWorkBook(clientWorkbookPath);

            excelApiLink.WriteCell(form.Path, form.ClientLine, 4, client);
            excelApiLink.WriteCell(form.Path, form.ClientLine + 1, 4, address);
            excelApiLink.WriteCell(form.Path, form.ClientLine + 2, 4, bp);
            excelApiLink.WriteCell(form.Path, form.ClientLine + 3, 4, postalCode);
            excelApiLink.WriteCell(form.Path, form.ClientLine + 3, 5, city);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Remplit les informations des étalons dans la page d'en-tête du formulaire
         * 
         */
        private void writeStandards(List<Standard> standards)
        {
            int linesToShift = standards.Count * 4;

            // Décalage des valeurs vers le bas
            excelApiLink.ShiftLines(form.Path, form.StandardLine, 1, 15, linesToShift);

            int standardLine = form.StandardLine;

            foreach (Standard standard in standards)
            {
                excelApiLink.WriteCell(form.Path, standardLine, 1, "Moyen:");
                excelApiLink.WriteCell(form.Path, standardLine, 4, standard.Name);

                excelApiLink.WriteCell(form.Path, standardLine + 1, 1, "Raccordement:");
                excelApiLink.WriteCell(form.Path, standardLine + 1, 4, standard.Raccordement);

                excelApiLink.MergeCells(form.Path, standardLine + 2, 4, standardLine + 2, 5);

                excelApiLink.WriteCell(form.Path, standardLine + 2, 1, "Validité:");
                excelApiLink.WriteCell(form.Path, standardLine + 2, 4, standard.Validity);

                standardLine += 4;
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Crée les feuilles de calculs nécessaires (délégué aux classes filles)
         * 
         */
        public abstract void CreateWorkSheets();

        /*-------------------------------------------------------------------------*/

        /**
         * Ecrit les valeurs des pièces dans le fichier excel (délégué aux classes filles)
         * 
         */
        public abstract void WritePiecesValues();

        /*-------------------------------------------------------------------------*/

        /**
         * Colle l'image de la signature sur le formulaire
         * 
         */
        private void signForm()
        {
            excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["HeaderPage"]);

            if (ConfigSingleton.Instance.Signature == null)
                throw new ConfigDataException("Il est impossible de signer le formulaire, la signature n'a pas été trouvée ou son format est incorrect");

            excelApiLink.PasteImage(form.Path, form.LineToSign, form.ColumnToSign, ConfigSingleton.Instance.Signature);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Exporte la première page du fichier excel en pdf
         * 
         */
        private void exportFirstPageToPdf()
        {
            excelApiLink.ExportFirstPageToPdf(form.Path, fileToSavePath.Replace(".xlsx", ".pdf"));
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Sauvegarde le fichier et ferme l'application
         * 
         */
        private void saveAndQuit()
        {
            excelApiLink.SaveWorkBook(form.Path, fileToSavePath);
            excelApiLink.CloseWorkBook(form.Path);
        }

        /*-------------------------------------------------------------------------*/
    }
}