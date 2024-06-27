using System.Text;
using System.IO;
using System.Globalization;

namespace Application.Parser
{
    /// <summary>
    /// Represents a text file parser that inherits from the base Parser class.
    /// This class is responsible for parsing a text file and returning a list of pieces.
    /// </summary>
    public class TextFileParser : Parser
    {
        private const string ENCODING = "iso-8859-1";
        private StreamReader? sr;
        private String fileToParse;
        int lineIndex = 1;

        // It allows to know if a header line is the first or not
        bool addPieceWhenHeaderMet = true;

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Initializes a new instance of the TextFileParser class.
        /// </summary>
        public TextFileParser()
        {
            this.fileToParse = "";
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Parses a text file and returns a list of pieces.
        /// </summary>
        /// <param name="fileToParse">The name of the file to parse.</param>
        /// <returns>The list of pieces.</returns>
        public override List<Data.Piece> ParseFile(string fileToParse)
        {
            this.fileToParse = fileToParse;
            base.dataParsed = [];

            sr = new StreamReader(fileToParse, Encoding.GetEncoding(ENCODING));

            string? line;

            // Read each line of the file
            while ((line = sr.ReadLine()) != null)
            {
                manageLineType(line);
                lineIndex++;
            }

            sr.Close();

            return dataParsed!;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Manages the line type based on its content.
        /// </summary>
        /// <param name="line">The line to analyze.</param>
        private void manageLineType(string line)
        {
            List<string> words;

            // Retrieves each word from the line in words by removing spaces
            words = [.. line.Split(' ')];
            words = words.Where((item, index) => item != "" && item != " ").ToList();

            // If the line is empty
            if (words.Count == 0) return;
            // If the line is a measure plan
            if (words[0][0] == '*') manageMeasurePlan(words);
            // If the line is a list of values from a measure
            else if (int.TryParse(words[0], out int testInt)) manageValueType(words);
            // If the line is a header
            else manageHeaderType(line);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Parses the header if a header line is detected.
        /// </summary>
        /// <param name="line">The line to analyze.</param>
        private void manageHeaderType(string line)
        {
            if (dataParsed == null) return;

            // If a new header is met
            if (this.addPieceWhenHeaderMet)
            {
                dataParsed.Add(new Data.Piece());
                this.addPieceWhenHeaderMet = false;
            }

            StringBuilder sb = new();
            sb.Append(line);

            this.lineIndex++;

            this.dataParsed[dataParsed.Count - 1].CreateHeader(sb.ToString());
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Retrieves the measure plan name from the line.
        /// </summary>
        /// <param name="words">The line to analyze as a list of words.</param>
        private void manageMeasurePlan(List<string> words)
        {
            if (!this.addPieceWhenHeaderMet) this.addPieceWhenHeaderMet = true;

            StringBuilder sb = new();
            sb.Append(words[0].AsSpan(5));

            for (int i = 1; i < words.Count; i++)
            {
                sb.Append(" " + words[i]);
            }

            string measurePlan = sb.ToString();
            dataParsed![^1].AddMeasurePlan(measurePlan);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Manages a value type line.
        /// </summary>
        /// <param name="words">The line to analyze as a list of words.</param>
        private void manageValueType(List<string> words)
        {
            if (!this.addPieceWhenHeaderMet) this.addPieceWhenHeaderMet = true;

            // Removes the unnecessary number that sometimes appears on certain measures
            if (int.TryParse(words[3], out int testInt)) words.RemoveAt(3);

            List<double> values = getLineToDoubleList(words, 3, words.Count - 1);

            string? nextLine = sr?.ReadLine();

            if (nextLine != null)
            {
                List<string> nextLineWords = [.. nextLine.Split(' ')];
                nextLineWords = nextLineWords.Where((item, index) => item != "" && item != " ").ToList();

                values.AddRange(getLineToDoubleList(nextLineWords, 0, nextLineWords.Count));
            }

            dataParsed![^1].AddData(getData(words, values));
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Converts a line of values in the texte file to a list of double values.
        /// </summary>
        /// <param name="words">The list of words in the line</param>
        /// <param name="startIndex">The first word to convert in the list</param>
        /// <param name="endIndex">The last word to convert in the list</param>
        /// <returns>A list of doubles converted from a line of the file</returns>
        private static List<double> getLineToDoubleList(List<string> words, int startIndex, int endIndex)
        {
            List<double> values = [];


            for (int i = startIndex; i < endIndex; i++)
            {
                if (double.TryParse(words[i], CultureInfo.InvariantCulture, out double testDouble))
                    values.Add(testDouble);
            }

            return values;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Returns the measure type of a line containing a measure.
        /// </summary>
        /// <param name="line">The line containing the measure.</param>
        /// <param name="values">The list of values.</param>
        /// <returns>The measure type of the line.</returns>
        private Data.Measure getData(List<string> line, List<double> values)
        {
            Data.Measure? data = Data.ConfigSingleton.Instance.GetData(line, values) ?? throw new Application.Exceptions.MeasureTypeNotFoundException(line[2], this.fileToParse, this.lineIndex);

            return data;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the file extension for the parser.
        /// </summary>
        /// <returns>The file extension.</returns>
        public override string GetFileExtension()
        {
            return "(*.mit;*.txt)|*.mit;*.txt";
        }

        /*-------------------------------------------------------------------------*/
    }
}
