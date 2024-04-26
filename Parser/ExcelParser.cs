using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Parser
{
    internal class ExcelParser : Parser
    {
        protected Excel.Application excelApp;
        protected Excel.Workbook? workbook;

        /*-------------------------------------------------------------------------*/

        public ExcelParser()
        {
            this.excelApp = new Excel.Application();

            base.header = new Dictionary<string, string>();

            base.header["Designation"] = "";
            base.header["N° de Plan"] = "";
            base.header["Client"] = "";
            base.header["Indice"] = "";
            base.header["Opérateurs"] = "";
            base.header["Observations"] = "";
        }

        /*-------------------------------------------------------------------------*/

        public override List<Data.Piece> ParseFile(string fileToParse)
        {
            base.dataParsed = new List<Data.Piece>();
            this.workbook = excelApp.Workbooks.Open(fileToParse);

            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];

            int row = 6;
            int col = 4;

            bool multiplePieces = worksheet.Cells[35, 1].Value == "Calcul";
            int nbPieces = 1;
            base.dataParsed.Add(new Data.Piece());

            // Vérifie s'il y a une ou plusieurs pièces
            if(multiplePieces)
            {
                nbPieces = this.getPiecesNumber(worksheet);

                for(int i = 1; i < nbPieces; i++)
                {
                    base.dataParsed.Add(new Data.Piece());
                }
            }

            String libelle = "";

            while (worksheet.Cells[row, col].Value != null)
            {
                libelle = worksheet.Cells[row + 7, col].Value;

                double nominalValue = worksheet.Cells[row, col].Value;
                double tolPlus = worksheet.Cells[row + 2, col].Value;
                double tolMinus = worksheet.Cells[row + 1, col].Value;

                Data.MeasureType? measureType = Data.ConfigSingleton.Instance.GetMeasureTypeFromLibelle(libelle);
                if(measureType == null)
                {
                    String cellName = worksheet.Cells[row + 7, col].Address;
                    this.workbook.Close();
                    this.excelApp.Quit();
                    throw new Exceptions.MeasureTypeNotFoundException(libelle, fileToParse, cellName);
                }

                String symbol = measureType.Symbol;

                // Pour chaque pièce (parcours de lignes)
                for(int i = 0; i < nbPieces; i++)
                {
                    int valueRow = multiplePieces ? 118 : 37;

                    Data.Data data = new Data.Data();
                    data.NominalValue = nominalValue;
                    data.TolerancePlus = tolPlus;
                    data.ToleranceMinus = tolMinus;
                    data.Symbol = symbol;
                    data.Value =  worksheet.Cells[valueRow + i, col].Value;

                    base.dataParsed[i].AddData(data);
                }

                col++;
            }

            this.workbook.Close();
            this.excelApp.Quit();

            return base.dataParsed;
        }

        /*-------------------------------------------------------------------------*/

        private int getPiecesNumber(Excel.Worksheet ws)
        {
            int row = 118;
            int nbPieces = 0;

            while (ws.Cells[row, 1].Value != null)
            {
                nbPieces = (int)ws.Cells[row, 1].Value;
                row++;
            }

            return nbPieces;
        }

        /*-------------------------------------------------------------------------*/
    }
}
