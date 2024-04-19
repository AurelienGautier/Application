using System.Text;
using System.IO;
using System.Globalization;

namespace Application
{
    internal class Parser
    {
        private const String ENCODING = "iso-8859-1";

        private List<Piece>? dataParsed;
        private StreamReader? sr;
        private Dictionary<String, String>? header;

        /*-------------------------------------------------------------------------*/

        /* ParseFile
         * 
         * Parse un fichier texte et retourne une liste de pièces
         * fileName : String - Nom du fichier à parser
         * return : List<Piece> - Liste de pièces
         * 
         */
        public List<Piece> ParseFile(String fileName)
        {
            this.dataParsed = new List<Piece>();
            this.sr = new StreamReader(fileName, Encoding.GetEncoding(ENCODING));

            String? line;

            try
            {
                while ((line = this.sr.ReadLine()) != null)
                {
                    this.manageLineType(line);
                }

                this.sr.Close();
            }
            catch
            {
                throw new Exceptions.IncorrectFormatException();
            }

            return this.dataParsed!;
        }

        /*-------------------------------------------------------------------------*/

        /* manageLineType
         * 
         * Gère le type de ligne en fonction de son contenu
         * line : String - Ligne à analyser
         *                                 
         */
        private void manageLineType(String line)
        {
            List<String> words;

            // Récupération de chaque mot de la ligne dans words en supprimant les espaces
            words = line.Split(' ').ToList();
            words = (words.Where((item, index) => item != "" && item != " ").ToArray()).ToList();

            if (words[0] == "Designation") this.manageHeaderType(line);
            else if (words[0][0] == '*') this.manageMeasurePlan(words);
            else this.manageValueType(words);
        }

        /*-------------------------------------------------------------------------*/

        /* manageHeaderType
         *
         * Parse l'en-tête si une ligne de type header est détectée
         * line : String - Ligne à analyser
         *                                 
         */
        private void manageHeaderType(String line)
        {
            this.dataParsed!.Add(new Piece());

            StringBuilder sb = new StringBuilder();
            sb.Append(line);

            for (int i = 0; i < 5; i++)
            {
                if (this.sr != null)
                    sb.Append('\n' + this.sr.ReadLine());
            }

            this.createHeader(sb.ToString());
        }

        /*-------------------------------------------------------------------------*/

        /* manageMeasurePlan
         *
         * Récupère le nom du plan de mesure de la ligne
         * words : List<String> - Ligne à analyser sous forme de liste des mots
         *
         */
        private void manageMeasurePlan(List<String> words)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(words[0].Substring(5));

            for (int i = 1; i < words.Count; i++)
            {
                sb.Append(" " + words[i]);
            }

            String measurePlan = sb.ToString();
            this.dataParsed![this.dataParsed!.Count - 1].AddMeasurePlan(measurePlan);
        }

        /*-------------------------------------------------------------------------*/

        /* manageValueType
         *
         * Gère une ligne de type valeur
         * words : List<String> - Ligne à analyser sous forme de liste des mots
         *
         */
        private void manageValueType(List<String> words)
        {
            this.dataParsed![this.dataParsed!.Count - 1].AddData(this.getData(words));

            List<double> values = new List<double>();

            // Suppression du nombre inutile qui apparaît parfois sur certaines mesures
            int testInt;
            if (int.TryParse(words[3], out testInt)) words.RemoveAt(3);

            for (int i = 3; i < words.Count - 1; i++)
            {
                values.Add(Convert.ToDouble(words[i].Replace('.', ',')));
            }

            String? nextLine = this.sr != null ? this.sr.ReadLine() : null;

            if (nextLine != null)
            {
                words = nextLine.Split(' ').ToList();
                words = (words.Where((item, index) => item != "" && item != " ").ToArray()).ToList();

                for (int i = 0; i < words.Count; i++)
                {
                    values.Add(Convert.ToDouble(words[i], new CultureInfo("en-US")));
                }
            }

            this.dataParsed![this.dataParsed!.Count - 1].SetValues(values);
        }

        /*-------------------------------------------------------------------------*/

        /* createHeader
         *
         * Crée l'en-tête à partir d'une chaîne de caractères
         * text : String - L'en-tête récupérée de manière brute depuis le fichier à analyser
         *
         */
        private void createHeader(string text)
        {
            this.header = new Dictionary<string, string>();

            string[] lines = text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

            foreach (var line in lines)
            {
                string[] parts = line.Split(new[] { ':' }, 3);

                string key = parts[0].Trim();
                string value = parts[2].Trim();

                this.header[key] = value;
            }

            string[] words = this.header["Opérateurs"].Split(' ');
            this.header["Opérateurs"] = words[1] + " " + words[0];
        }

        /*-------------------------------------------------------------------------*/

        /* getData
         *
         * Retourne le type de mesure d'une ligne contenant une mesure
         * line : List<String> - Ligne contenant la mesure
         * return : Data - Type de mesure de la ligne
         *
         */
        private Data.Data getData(List<String> line)
        {
            String symbol = "";

            if (line[2] == "Distance" || line[2] == "Diameter" || line[2] == "Pos." || line[2] == "Angle" || line[2] == "Result" || line[2] == "Min.Ax/2")
            {
                if (line[2] == "Diameter") symbol = "⌀";

                if (line[2] == "Pos.")
                {
                    symbol = line[3];
                    line[2] += line[3];
                    line.RemoveAt(3);
                }

                return new Data.Data(symbol);
            }
            else if (line[2] == "Ax:R/Out" || line[2] == "CirR/Out" || line[2] == "Symmetry")
                return new Data.DataAxCirOut("");
            else if (line[2] == "Concentr") return new Data.DataConcentricity("");
            else if (line[2] == "Position")
            {
                symbol = "⊕";
                return new Data.DataPosition(symbol);
            }
            else if (line[2] == "Flatness") symbol = "⏥";
            else if (line[2] == "Rectang.") symbol = "_";
            else if (line[2] == "Parallel") symbol = "//";

            return new Data.DataSimple(symbol);
        }

        /*-------------------------------------------------------------------------*/

        /* GetHeader
         * 
         * Retourne l'en-tête du fichier
         * 
         */
        public Dictionary<String, String> GetHeader()
        {
            return this.header!;
        }

        /*-------------------------------------------------------------------------*/
    }
}
