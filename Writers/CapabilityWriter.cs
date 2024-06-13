using Application.Data;

namespace Application.Writers
{
    class CapabilityWriter : ExcelWriter
    {
        private readonly int maxLines = 25;

        public CapabilityWriter(String fileName, Form form) : base(fileName, form)
        {
        }

        /// <summary>
        /// Creates a worksheet for each capability measure
        /// </summary>
        /// <exception cref="Exceptions.IncorrectValuesToTreatException"></exception>
        public override void CreateWorkSheets()
        {
            if(!base.form.Modify)
            {
                if(form.CapabilityMeasureNumber == null)
                {
                    throw new Exceptions.IncorrectValuesToTreatException("Le nombre de mesures de capacité n'a pas été renseigné.");
                }

                Console.WriteLine(form.CapabilityMeasureNumber.Count);

                for(int i = 2; i <= form.CapabilityMeasureNumber.Count; i++)
                {
                    excelApiLink.CopyWorkSheet(form.Path, "Capa", "Capa (" + i + ")");
                }
            }
        }

        /// <summary>
        /// Writes the values of each capability measure in a different worksheet
        /// </summary>
        /// <exception cref="Exceptions.IncoherentValueException"></exception>
        public override void WritePiecesValues()
        {
            excelApiLink.ChangeWorkSheet(form.Path, "Capa");
            int linesWritten = 0;
            int line = form.FirstLine;
            int col = 5;

            if (form.CapabilityMeasureNumber == null) return;
            List<int> capabilityMeasureNumber = form.CapabilityMeasureNumber;

            // Write the values of the pieces in the capability form
            for (int i = 0; i < capabilityMeasureNumber.Count; i++)
            {
                if(i > 0)
                {
                    excelApiLink.ChangeWorkSheet(form.Path, "Capa (" + (i + 1) + ")");
                    line = form.FirstLine;
                    col = 5;
                    linesWritten = 0;
                }

                int num = capabilityMeasureNumber[i];
                foreach (Piece piece in pieces)
                {
                    try
                    {
                        double currentValue = piece.GetData()[0][num].Value;

                        excelApiLink.WriteCell(form.Path, line, col, currentValue);
                        linesWritten++;
                        line++;

                        if (linesWritten == maxLines)
                        {
                            linesWritten = 0;
                            line = form.FirstLine;
                            col++;
                        }
                    }
                    catch
                    {
                        throw new Exceptions.IncoherentValueException("Le format du fichier n'est pas cohérent avec la valeur l'un des numéros de mesure fournis.");
                    }
                }
            }
        }
    }
}
