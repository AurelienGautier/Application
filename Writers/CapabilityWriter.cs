using Application.Data;

namespace Application.Writers
{
    class CapabilityWriter : ExcelWriter
    {
        private readonly int maxLines = 25;

        public CapabilityWriter(String fileName, Form form) : base(fileName, form)
        {
        }

        public override void CreateWorkSheets()
        {
            // The capability form only has one measure page
        }

        public override void WritePiecesValues()
        {
            excelApiLink.ChangeWorkSheet(form.Path, "Capa");
            int linesWritten = 0;
            int line = form.FirstLine;
            int col = 5;

            if(form.CapabilityMeasureNumber == null) return;
            int capabilityMeasureNumber = (int)form.CapabilityMeasureNumber;

            // Write the values of the pieces in the capability form
            foreach (Piece piece in pieces)
            {
                try
                {
                    double currentValue = piece.GetData()[0][capabilityMeasureNumber].Value;

                    excelApiLink.WriteCell(form.Path, line, col, currentValue);
                    linesWritten++;
                    line++;

                    if (linesWritten == maxLines)
                    {
                        linesWritten = 0;
                        line = form.FirstLine;
                        col ++;
                    }
                }
                catch
                {
                    throw new Exceptions.IncoherentValueException("Le format du fichier n'est pas cohérent avec la valeur le numéro de mesure fourni.");
                }
            }
        }
    }
}
