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
            int nbLine = 55;

            while ((line = this.sr.ReadLine()) != null)
            {
                words = line.Split(' ').ToList();
                words = ((String[])words.Where((item, index) => item != "" && item != " ").ToArray()).ToList();

                LineType type = this.GetLineType(words);

                this.ManageLineType(type, words);

                nbLine++;
            }

            this.sr.Close();

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
                        Console.WriteLine(words);
                    this.dataParsed[this.dataParsed.Count - 1].AddData(this.GetData(words));

                    List<double> values = new List<double>();

                    for(int i = 3; i < words.Count - 1; i++)
                    {
                        values.Add(Convert.ToDouble(words[i], new CultureInfo("en-US")));
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
            if (line.Count == 0) return LineType.VOID;
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

            if (line[2] == "Concentr") return new Data.DataConcentricity();
            if (line[2] == "Roundnes") return new Data.DataRoundNess();
            if (line[2] == "Symmetry") return new Data.DataSymmetry();
            if (line[2] == "Rectang.") return new Data.DataRectangle();
            if (line[2] == "Position") return new Data.DataPosition();
            if (line[2] == "Flatness") return new Data.DataFlatness();

            return new Data.Data();
        }
    }
}
