using Application.Writers;

namespace Application.Parser
{
    public class ExcelParser : Parser
    {
        public ExcelParser()
        {
        }

        /*-------------------------------------------------------------------------*/

        public override List<Data.Piece> ParseFile(string fileToParse)
        {
            base.dataParsed = new List<Data.Piece>();
            ExcelLibraryLinkSingleton excelApiLink = ExcelLibraryLinkSingleton.Instance;
            excelApiLink.OpenWorkBook(fileToParse);
            excelApiLink.ChangeWorkSheet(fileToParse, 1);

            int row = 6;
            int col = 4;

            bool multiplePieces = excelApiLink.ReadCell(fileToParse, 35, 1) == "Calcul";
            int nbPieces = 1;
            base.dataParsed.Add(new Data.Piece());

            // Vérifie s'il y a une ou plusieurs pièces
            if(multiplePieces)
            {
                nbPieces = this.getPiecesNumber(fileToParse);

                for(int i = 1; i < nbPieces; i++)
                {
                    base.dataParsed.Add(new Data.Piece());
                }
            }

            String libelle = "";

            while(excelApiLink.ReadCell(fileToParse, row, col) != "")
            {
                libelle = excelApiLink.ReadCell(fileToParse, row + 7, col);

                double nominalValue = Convert.ToDouble(excelApiLink.ReadCell(fileToParse, row, col));
                double tolPlus = Convert.ToDouble(excelApiLink.ReadCell(fileToParse, row + 2, col));
                double tolMinus = Convert.ToDouble(excelApiLink.ReadCell(fileToParse, row + 1, col));

                Data.MeasureType? measureType = Data.ConfigSingleton.Instance.GetMeasureTypeFromLibelle(libelle);
                if(measureType == null)
                {
                    String cellName = excelApiLink.GetCellAddress(row + 7, col);
                    excelApiLink.CloseWorkBook(fileToParse);
                    throw new Exceptions.MeasureTypeNotFoundException(libelle, fileToParse, cellName);
                }

                // Pour chaque pièce (parcours de lignes)
                for(int i = 0; i < nbPieces; i++)
                {
                    int valueRow = multiplePieces ? 118 : 37;

                    Data.Measure data = new Data.Measure(measureType);
                    data.NominalValue = nominalValue;
                    data.TolerancePlus = tolPlus;
                    data.ToleranceMinus = tolMinus;
                    data.Value = Convert.ToDouble(excelApiLink.ReadCell(fileToParse, valueRow + i, col));

                    base.dataParsed[i].AddData(data);
                }

                col++;
            }

            excelApiLink.CloseWorkBook(fileToParse);

            return base.dataParsed;
        }

        /*-------------------------------------------------------------------------*/

        private int getPiecesNumber(String fileToParse)
        {
            int row = 118;
            int nbPieces = 0;

            while (ExcelLibraryLinkSingleton.Instance.ReadCell(fileToParse, row, 1) != "")
            {
                nbPieces++;

                row++;
            }

            return nbPieces;
        }

        /*-------------------------------------------------------------------------*/

        public override string GetFileExtension()
        {
            return "(*.xlsx;*.xlsm)|*.xlsx;*.xlsm";
        }

        /*-------------------------------------------------------------------------*/
    }
}
