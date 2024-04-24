using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    internal class FivePiecesWriter : ExcelWriter
    {
        private int pageNumber;
        private readonly List<List<String>> measurePlans;
        private readonly List<List<List<Data.Data>>> pieceData;
        private int linesWritten;
        private Excel.Worksheet ws;
        private int min;
        private int max;
        private const int MAX_LINES = 23;

        /**
         * FivePiecesWriter
         * 
         * Constructeur de la classe
         * fileName : string - Nom du fichier à sauvegarder
         * 
         */
        public FivePiecesWriter(string fileName) : base(fileName, 17, 1, Environment.CurrentDirectory + "\\form\\rapport5pieces")
        {
            this.pageNumber = 1;
            this.measurePlans = new List<List<String>>();
            this.pieceData = new List<List<List<Data.Data>>>();
            this.linesWritten = 0;
            this.ws = base.workbook.Sheets["Mesures"];
            this.min = 0;
            this.max = 5;
        }

        /**
         * CreateWorkSheets
         * 
         * Crée toutes les pages Excel nécessaire pour insérer toutes les données
         * 
         */
        public override void CreateWorkSheets()
        {
            int TotalPageNumber = pieces[0].GetLinesToWriteNumber() / MAX_LINES + 1;

            int iterations = base.pieces.Count / 5;
            if (base.pieces.Count % 5 != 0) iterations++;

            TotalPageNumber *= iterations;

            for(int i = 4; i <= TotalPageNumber; i++)
            {
                workbook.Sheets["Mesures"].Copy(Type.Missing, workbook.Sheets[workbook.Sheets.Count]);
            }
        }

        /**
         * WritePiecesValues
         * 
         * Écrit les valeurs de mesure des pièces dans le fichier Excel
         * 
         */
        public override void WritePiecesValues()
        {
            for(int i = 0; i < base.pieces.Count; i++)
            {
                this.measurePlans.Add(base.pieces[i].GetMeasurePlans());
                this.pieceData.Add(base.pieces[i].GetData());
            }

            int iterations = pieceData.Count / 5;
            if(pieceData.Count % 5 != 0) iterations++;

            for (int i = 0; i < iterations; i++)
            {
                this.Write5pieces();

                this.min += 5;

                if (i == pieceData.Count / 5 - 1 && pieceData.Count % 5 != 0) this.max = pieceData.Count;
                else this.max += 5;

                if(i < iterations - 1) this.ChangePage();
            }
        }

        /**
         * Write5pieces
         * 
         * Écrit toutes les valeurs pour un groupe de 5 pièces dans le fichier Excel
         * 
         */
        public void Write5pieces()
        {
            // Pour chaque plan
            for (int i = 0; i < pieceData[0].Count; i++)
            {
                // Écriture du plan
                if (measurePlans[0][i] != "")
                {
                    ws.Cells[base.currentLine, base.currentColumn].Value = measurePlans[0][i];
                    base.currentLine++;
                    this.linesWritten++;
                }

                // Changement de page si l'actuelle est complète
                if (this.linesWritten == MAX_LINES) { this.ChangePage(); }

                // Pour chaque mesure du plan
                for (int j = 0; j < pieceData[0][i].Count; j++)
                {
                    ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[0][i][j].GetNominalValue();
                    ws.Cells[base.currentLine, base.currentColumn + 2].Value = pieceData[0][i][j].GetTolPlus();
                    ws.Cells[base.currentLine, base.currentColumn + 3].Value = pieceData[0][i][j].GetTolMinus();

                    base.currentColumn += 3;

                    // Écriture des valeurs des pièces
                    for(int k = this.min; k < this.max; k++)
                    {
                        base.currentColumn += 3;
                        ws.Cells[base.currentLine, base.currentColumn].Value = pieceData[k][i][j].GetValue();
                    }

                    base.currentColumn -= (3 + 3 * (this.max - this.min));

                    base.currentLine++;
                    this.linesWritten++;

                    // Changement de page si l'actuelle est complète
                    if (this.linesWritten == MAX_LINES) { this.ChangePage(); }
                }
            }
        }

        /**
         * ChangePage
         * 
         * Passe à la page de mesure suivante
         * 
         */
        public void ChangePage()
        {
            this.pageNumber++;

            ws = base.workbook.Sheets["Mesures (" + this.pageNumber.ToString() + ")"];

            base.currentLine = 17;
            this.linesWritten = 0;

            int col = 7;

            for (int i = this.min; i < this.min + 5; i++)
            {
                ws.Cells[15, col].Value = i + 1;
                col += 3;
            }
        }

        /**
         * WriteHeader
         * 
         * Remplit l'entête du rapport Excel
         * 
         * header : Dictionary<string, string> - Dictionnaire contenant les informations de l'entête
         * designLine : int - Numéro de la ligne où écrire la désignation
         * 
         */
        public void WriteHeader(Dictionary<string, string> header, int designLine)
        {
            Excel.Worksheet worksheet = base.workbook.Sheets["Rapport d'essai dimensionnel"];

            worksheet.Cells[designLine, 4] = header["Designation"];
            worksheet.Cells[designLine + 2, 4] = header["N° de Plan"];
            worksheet.Cells[designLine + 4, 4] = header["Indice"];
        }

        /**
         * ExcportFirstPageToPdf
         * 
         * Exporte la première page du formulaire Excel en PDF
         * 
         */
        public override void ExportFirstPageToPdf()
        {
        }
    }
}
