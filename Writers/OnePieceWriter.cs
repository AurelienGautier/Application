using Excel = Microsoft.Office.Interop.Excel;
using Application.Data;
using Microsoft.Office.Interop.Excel;

namespace Application.Writers
{
    internal class OnePieceWriter : ExcelWriter
    {
        private const int MAX_LINES = 22;
        private int linesWritten;
        private int pageNumber;

        /*-------------------------------------------------------------------------*/

        public OnePieceWriter(string fileName, Form form) : base(fileName, form)
        {
            this.linesWritten = 0;
            this.pageNumber = 1;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Crée suffisamment de pages Excel pour écrire les données de la pièce.
         * 
         * La première feuille est la feuille "Mesures" qui contient les données de la pièce.
         * 
         * Si le nombre de lignes à écrire est supérieur à MAX_LINES, des copies de la feuille "Mesures" sont créées.
         * 
         */
        public override void CreateWorkSheets()
        {
            int linesToWrite = pieces[0].GetLinesToWriteNumber();

            int numberOfPages = linesToWrite / MAX_LINES;

            for (int i = 2; i < numberOfPages; i++)
            {
                excelApiLink.CopyWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"], ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + i.ToString() + ")");
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Écrit les valeurs de mesure des pièces dans les feuilles Excel.
         * 
         */
        public override void WritePiecesValues()
        {
            excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"]);

            List<String> measurePlans = pieces[0].GetMeasurePlans();
            List<List<Data.Measure>> pieceData = pieces[0].GetData();


            for (int i = 0; i < pieceData.Count; i++)
            {
                // Écriture du plan
                if (measurePlans[i] != "")
                {
                    if (this.isLastLine(pieceData, i, 0) && this.isNextLineEmpty())
                        this.throwIncoherentValueException();
                        
                    excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 1, measurePlans[i]);
                    base.currentLine++;
                    this.linesWritten++;
                }

                // Changement de page si l'actuelle est complète
                if (this.linesWritten == MAX_LINES)
                {
                    this.ChangePage();
                }
                
                // Écriture des données ligne par ligne
                for (int j = 0; j < pieceData[i].Count; j++)
                {
                    if(!base.form.Modify)
                    {
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 1, pieceData[i][j].MeasureType.Symbol);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 2, pieceData[i][j].NominalValue);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 4, pieceData[i][j].TolerancePlus);
                        excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 5, pieceData[i][j].ToleranceMinus);
                    }

                    // Throws an error if the number of measures in the report is different from the number of measures in the source file
                    if (form.Modify) 
                    {
                        if (excelApiLink.ReadCell(form.Path, base.currentLine, base.currentColumn + 2) == "")
                            this.throwIncoherentValueException();
                        else if (this.isLastLine(pieceData, i, j) && this.isNextLineEmpty())
                            this.throwIncoherentValueException();
                    }

                    excelApiLink.WriteCell(form.Path, base.currentLine, base.currentColumn + 6, pieceData[i][j].Value);

                    base.currentLine++;
                    this.linesWritten++;

                    if (this.linesWritten == MAX_LINES)
                    {
                        this.ChangePage();
                    }
                }
            }
        }

        /*-------------------------------------------------------------------------*/

        private void ChangePage()
        {
            pageNumber++;

            try
            {
                excelApiLink.ChangeWorkSheet(form.Path, ConfigSingleton.Instance.GetPageNames()["MeasurePage"] + " (" + pageNumber.ToString() + ")");
            }
            catch
            {
                excelApiLink.CloseWorkBook(form.Path);

                throw new Exceptions.IncoherentValueException("Le nombre de mesures n'est pas le même entre le rapport à modifier et le ou les fichiers sources");
            }

            base.currentLine -= linesWritten;
            linesWritten = 0;
        }

        /*-------------------------------------------------------------------------*/

        private bool isNextLineEmpty()
        {
            if (excelApiLink.ReadCell(form.Path, base.currentLine + 1, base.currentColumn + 2) != "")
                return true;

            return false;
        }

        /*-------------------------------------------------------------------------*/

        private bool isLastLine(List<List<Data.Measure>> pieceData, int i, int j)
        {
            if (i != pieceData.Count - 1) return false; 

            if (pieceData[i].Count == 0 || j == pieceData[i].Count - 1)
                return true;
            
            return false;
        }

        /*-------------------------------------------------------------------------*/

        private void throwIncoherentValueException()
        {
            excelApiLink.CloseWorkBook(form.Path);

            throw new Exceptions.IncoherentValueException("Le nombre de mesures n'est pas le même entre le rapport à modifier et le ou les fichiers sources");
        }

        /*-------------------------------------------------------------------------*/
    }
}
