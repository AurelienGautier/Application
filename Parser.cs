﻿using System;
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
        private List<List<String>> linesParsed;
        private List<Piece> dataParsed;
        private StreamReader sr;

        [DllImport("Kernel32")]
        public static extern void AllocConsole();

        [DllImport("Kernel32", SetLastError = true)]
        public static extern void FreeConsole();

        public Parser()
        {
            this.linesParsed = new List<List<String>>();
            this.dataParsed = new List<Piece>();
            this.sr = new StreamReader("C:\\Users\\LaboTri-PC2\\Desktop\\dev\\test\\test.txt");

            AllocConsole();
        }

        ~Parser() 
        {
            this.sr.Close();

            FreeConsole();
        }

        public List<List<String>> ParseFile()
        {
            /*try
            {*/
                String? line;
                List<String> words;
                int nbLine = 55;

                while ((line = this.sr.ReadLine()) != null)
                {
                    words = line.Split(' ').ToList();
                    words = ((String[])words.Where((item, index) => item != "" && item != " ").ToArray()).ToList();

                    this.linesParsed.Add(words);

                    LineType type = this.GetLineType(words);

                    this.ManageLineType(type, words);

                    nbLine++;
                }

                this.PrintAll();
            /*}
            catch
            {
                Console.WriteLine("Erreur : le fichier n'a pas pu être parsé correctement.");
                this.sr.Close();
            }*/

            return this.linesParsed;
        }

        public void PrintAll()
        {
            foreach(Piece piece in this.dataParsed)
            {
                piece.PrintTrucs();
            }
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

                    this.dataParsed[this.dataParsed.Count - 1].setValues(values);

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
            if (line[2] == "Diameter") return new Data.DataDiamater();
            if (line[2] == "Concentr") return new Data.DataConcentricity();
            if (line[2] == "Distance") return new Data.DataDistance();
            if (line[2] == "Roundnes") return new Data.DataRoundNess();
            if (line[2] == "Symmetry") return new Data.DataSymmetry();
            if (line[2] == "Pos.")
            {
                line[2] += line[3];
                line.RemoveAt(3);
                return new Data.DataPosX();
            }

            return null;
        }
    }
}
