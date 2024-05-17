using Application.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Data
{
    internal class Header
    {
        public String Designation { get; set; }
        public String PlanNb { get; set; }
        public String Index { get; set; }
        public String ClientName { get; set; }

        public Header()
        {
            this.Designation = "";
            this.PlanNb = "";
            this.Index = "";
            this.ClientName = "";
        }

        public void FillHeader(Dictionary<String, String> rawHeader)
        {
            Dictionary<String, String>? matchWithFileFields = ConfigSingleton.Instance.GetHeaderFieldsMatch();

            if (matchWithFileFields == null)
                throw new ConfigDataException("Le fichier de configuration contenant les paramètres de l'en-tête est incorrect ou introuvable");

            if (rawHeader.ContainsKey(matchWithFileFields["Designation"]))
                this.Designation = rawHeader[matchWithFileFields["Designation"]];

            if (rawHeader.ContainsKey(matchWithFileFields["PlanNb"]))
                this.PlanNb = rawHeader[matchWithFileFields["PlanNb"]];

            if (rawHeader.ContainsKey(matchWithFileFields["Index"]))
                this.Index = rawHeader[matchWithFileFields["Index"]];
            
            if (rawHeader.ContainsKey(matchWithFileFields["ClientName"]))
                this.ClientName = rawHeader[matchWithFileFields["ClientName"]];
        }   
    }
}
