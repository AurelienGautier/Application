using Application.Exceptions;

namespace Application.Data
{
    internal class Header
    {
        public String Designation { get; set; }
        public String PlanNb { get; set; }
        public String Index { get; set; }
        public String ClientName { get; set; }
        public String ObservationNum { get; set; }
        public String PieceReceptionDate { get; set; }
        public String Observations { get; set; }

        public Header()
        {
            this.Designation = "";
            this.PlanNb = "";
            this.Index = "";
            this.ClientName = "";
            this.ObservationNum = "";
            this.PieceReceptionDate = "";
            this.Observations = "";
        }

        public void FillHeader(Dictionary<String, String> rawHeader)
        {
            Dictionary<String, String>? matchWithFileFields = ConfigSingleton.Instance.GetHeaderFieldsMatch();

            if (matchWithFileFields == null)
                throw new ConfigDataException("Le fichier de configuration contenant les paramètres de l'en-tête est incorrect ou introuvable");

            try
            {
                if (rawHeader.ContainsKey(matchWithFileFields["Designation"]))
                    this.Designation = rawHeader[matchWithFileFields["Designation"]];

                if (rawHeader.ContainsKey(matchWithFileFields["PlanNb"]))
                    this.PlanNb = rawHeader[matchWithFileFields["PlanNb"]];

                if (rawHeader.ContainsKey(matchWithFileFields["Index"]))
                    this.Index = rawHeader[matchWithFileFields["Index"]];
            
                if (rawHeader.ContainsKey(matchWithFileFields["ClientName"]))
                    this.ClientName = rawHeader[matchWithFileFields["ClientName"]];

                if (rawHeader.ContainsKey(matchWithFileFields["ObservationNum"]))
                    this.ObservationNum = rawHeader[matchWithFileFields["ObservationNum"]];

                if (rawHeader.ContainsKey(matchWithFileFields["PieceReceptionDate"]))
                    this.PieceReceptionDate = rawHeader[matchWithFileFields["PieceReceptionDate"]];

                if (rawHeader.ContainsKey(matchWithFileFields["Observations"]))
                    this.Observations = rawHeader[matchWithFileFields["Observations"]];
            }
            catch
            {
                // If the file is not correctly formatted, the fields are left empty
            }
        }   
    }
}
