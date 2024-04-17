using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections.Generic;
using System.Windows.Shapes;
using System.Globalization;
using Application.Data;

namespace Application
{
    enum LineType
    {
        HEADER,
        MEASURE_TYPE,
        VALUE,
        VOID
    }

    internal class Parser
    {
        private List<Piece>? dataParsed;
        private StreamReader? sr;
        private Dictionary<String, String>? header;

        public void CreateHeader(string text)
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

        public Dictionary<String, String> GetHeader()
        {
            return this.header!;
        }

        public List<Piece> ParseFile(String fileName)
        {
            this.dataParsed = new List<Piece>();
            this.sr = new StreamReader(fileName, Encoding.GetEncoding("iso-8859-1"));

            String? line;
            List<String> words;

            try
            {
                while ((line = this.sr.ReadLine()) != null)
                {
                    words = line.Split(' ').ToList();
                    words = (words.Where((item, index) => item != "" && item != " ").ToArray()).ToList();

                    LineType type = this.GetLineType(words);

                    this.ManageLineType(type, words, line);
                }

                this.sr.Close();
            }
            catch
            {
                throw new Exceptions.IncorrectFormatException();
            }

            return this.dataParsed!;
        }

        public void ManageLineType(LineType type, List<String> words, String line)
        {
            switch (type)
            {
                case LineType.HEADER:
                    {
                        this.dataParsed!.Add(new Piece());

                        StringBuilder sb = new StringBuilder();
                        sb.Append(line);

                        for (int i = 0; i < 5; i++)
                        {
                            if(this.sr != null)
                                sb.Append('\n' + this.sr.ReadLine());
                        }

                        this.CreateHeader(sb.ToString());

                        break;
                    }
                case LineType.MEASURE_TYPE:
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append(words[0].Substring(5));

                        for (int i = 1; i < words.Count; i++)
                        {
                            sb.Append(" " + words[i]);
                        }

                        String measureType = sb.ToString();
                        this.dataParsed![this.dataParsed!.Count - 1].AddMeasureType(measureType);

                        break;
                    }
                case LineType.VALUE:
                    {
                        this.dataParsed![this.dataParsed!.Count - 1].AddData(this.GetData(words));

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

                        break;
                    }
            }
        }

        public LineType GetLineType(List<String> line)
        {
            if (line.Count == 0) return LineType.VOID;
            if (line[0] == "Designation") return LineType.HEADER;
            if (line[0][0] == '*') return LineType.MEASURE_TYPE;
            return LineType.VALUE;
        }

        public Data.Data GetData(List<String> line)
        {
            if (line[2] == "Distance" || line[2] == "Diameter" || line[2] == "Pos." || line[2] == "Angle" || line[2] == "Result" || line[2] == "Min.Ax/2")
            {
                if (line[2] == "Pos.")
                {
                    line[2] += line[3];
                    line.RemoveAt(3);
                }

                return new Data.Data();
            }

            if (line[2] == "Ax:R/Out" || line[2] == "CirR/Out" || line[2] == "Symmetry")
                return new Data.DataAxCirOut();

            if (line[2] == "Concentr") return new Data.DataConcentricity();
            if (line[2] == "Position") return new Data.DataPosition();

            return new Data.DataSimple();
        }
    }
}
