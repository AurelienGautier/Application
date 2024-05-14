using System.Text;
using System.IO;
using System.Globalization;

namespace Application.Parser
{
    internal class TextFileParser : Parser
    {
        private const string ENCODING = "iso-8859-1";
        private StreamReader? sr;
        private String fileToParse;
        int lineIndex = 1;

        /*-------------------------------------------------------------------------*/

        public TextFileParser()
        {
            this.fileToParse = "";
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Parse un fichier texte et retourne une liste de pièces
         * fileToParse : String - Nom du fichier à parser
         * return : List<Data.Piece> - Liste de pièces
         * 
         */
        public override List<Data.Piece> ParseFile(string fileToParse)
        {
            this.fileToParse = fileToParse;
            base.dataParsed = new List<Data.Piece>();

            sr = new StreamReader(fileToParse, Encoding.GetEncoding(ENCODING));

            string? line;

            while ((line = sr.ReadLine()) != null)
            {
                manageLineType(line);
                lineIndex++;
            }

            sr.Close();

            return dataParsed!;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Gère le type de ligne en fonction de son contenu
         * line : String - Ligne à analyser
         *                                 
         */
        private void manageLineType(string line)
        {
            List<string> words;

            // Récupération de chaque mot de la ligne dans words en supprimant les espaces
            words = line.Split(' ').ToList();
            words = words.Where((item, index) => item != "" && item != " ").ToList();

            if(words.Count == 0) return;
            if (words[0] == "Designation") manageHeaderType(line);
            else if (words[0][0] == '*') manageMeasurePlan(words);
            else manageValueType(words);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Parse l'en-tête si une ligne de type header est détectée
         * line : String - Ligne à analyser
         *                                 
         */
        private void manageHeaderType(string line)
        {
            if (dataParsed == null) return;

            dataParsed.Add(new Data.Piece());

            StringBuilder sb = new StringBuilder();
            sb.Append(line);

            for (int i = 0; i < 5; i++)
            {
                if (sr != null)
                    sb.Append('\n' + sr.ReadLine());
            }

            this.lineIndex += 5;

            this.dataParsed[dataParsed.Count - 1].CreateHeader(sb.ToString());
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Récupère le nom du plan de mesure de la ligne
         * words : List<String> - Ligne à analyser sous forme de liste des mots
         *
         */
        private void manageMeasurePlan(List<string> words)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(words[0].Substring(5));

            for (int i = 1; i < words.Count; i++)
            {
                sb.Append(" " + words[i]);
            }

            string measurePlan = sb.ToString();
            dataParsed![dataParsed!.Count - 1].AddMeasurePlan(measurePlan);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Gère une ligne de type valeur
         * words : List<String> - Ligne à analyser sous forme de liste des mots
         *
         */
        private void manageValueType(List<string> words)
        {
            // Suppression du nombre inutile qui apparaît parfois sur certaines mesures
            int testInt;
            if (int.TryParse(words[3], out testInt)) words.RemoveAt(3);

            List<double> values = getLineToDoubleList(words, 3, words.Count - 1);

            string? nextLine = sr != null ? sr.ReadLine() : null;

            if (nextLine != null)
            {
                List<string> nextLineWords = nextLine.Split(' ').ToList();
                nextLineWords = nextLineWords.Where((item, index) => item != "" && item != " ").ToList();

                values.AddRange(getLineToDoubleList(nextLineWords, 0, nextLineWords.Count));
            }

            dataParsed![dataParsed!.Count - 1].AddData(getData(words, values));
        }

        /*-------------------------------------------------------------------------*/

        private List<double> getLineToDoubleList(List<string> words, int startIndex, int endIndex)
        {
            List<double> values = new List<double>();

            double testDouble;

            for (int i = startIndex; i < endIndex; i++)
            {
                if (double.TryParse(words[i].Replace('.', ','), out testDouble))
                    values.Add(Convert.ToDouble(words[i].Replace('.', ',')));
            }

            return values;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Retourne le type de mesure d'une ligne contenant une mesure
         * line : List<String> - Ligne contenant la mesure
         * return : Data - Type de mesure de la ligne
         *
         */
        private Data.Data getData(List<string> line, List<double> values) 
        {
            Data.Data? data = Data.ConfigSingleton.Instance.GetData(line, values);

            if (data == null)
            {
                throw new Application.Exceptions.MeasureTypeNotFoundException(line[2], this.fileToParse, this.lineIndex);
            }

            return data;
        }

        /*-------------------------------------------------------------------------*/

        public override string GetFileExtension()
        {
            return "(*.mit;*.txt)|*.mit;*.txt";
        }

        /*-------------------------------------------------------------------------*/
    }
}
