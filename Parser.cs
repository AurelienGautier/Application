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
        private List<Piece> dataParsed;
        private StreamReader sr;

        public Parser(String fileName)
        {
            this.dataParsed = new List<Piece>();
            this.sr = new StreamReader(fileName);
        }

        public List<Piece> ParseFile()
        {
            String? line;
            List<String> words;

            try
            {
                while ((line = this.sr.ReadLine()) != null)
                {
                    words = line.Split(' ').ToList();
                    words = ((String[])words.Where((item, index) => item != "" && item != " ").ToArray()).ToList();

                    LineType type = this.GetLineType(words);

                    this.ManageLineType(type, words);
                }

                this.sr.Close();
            }
            catch
            {
                throw new Exceptions.IncorrectFormatException();
            }

            return this.dataParsed;
        }

        public void ManageLineType(LineType type, List<String> words)
        {
            switch(type)
            {
                case LineType.HEADER:
                {
                    this.dataParsed.Add(new Piece());

                    for (int i = 0; i < 5; i++)
                    {
                        this.sr.ReadLine();
                    }

                    break;
                }
                case LineType.MEASURE_TYPE:
                {
                    String measureType = words[0].Substring(5);

                    for(int i = 1; i < words.Count; i++)
                    {
                        measureType += " " + words[i];
                    }

                    this.dataParsed[this.dataParsed.Count - 1].AddMeasureType(measureType);

                    break;
                }
                case LineType.VALUE:
                {
                    this.dataParsed.Last().AddData(this.GetData(words));

                    List<double> values = new List<double>();

                    // Suppression du nombre inutile qui apparaît parfois sur certaines mesures
                    int testInt;
                    if (int.TryParse(words[3], out testInt)) words.RemoveAt(3);

                    for(int i = 3; i < words.Count - 1; i++)
                    {
                        values.Add(Convert.ToDouble(words[i].Replace('.', ',')));
                    }

                    String? nextLine = this.sr.ReadLine();

                    if(nextLine != null)
                    {
                        words = nextLine.Split(' ').ToList();
                        words = ((String[])words.Where((item, index) => item != "" && item != " ").ToArray()).ToList();

                        for(int i = 0; i < words.Count; i++)
                        {
                            values.Add(Convert.ToDouble(words[i], new CultureInfo("en-US")));
                        }
                    }

                    this.dataParsed[this.dataParsed.Count - 1].SetValues(values);

                    break;
                }
            }
        }

        public LineType GetLineType(List<String> line)
        {
            if(line.Count == 0) return LineType.VOID;
            if (line[0] == "Designation") return LineType.HEADER;
            if (line[0][0] == '*') return LineType.MEASURE_TYPE;
            return LineType.VALUE;
        }

        public Data.Data GetData(List<String> line)
        {
            if (line[2] == "Distance" || line[2] == "Diameter" || line[2] == "Pos." || line[2] == "Angle" || line[2] == "Result")
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

            return new Data.DataSimple();
        }
    }
}
