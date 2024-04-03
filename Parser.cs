using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections.Generic;

namespace Application
{

    internal class Parser
    {
        private List<String[]> LinesParsed;

        [DllImport("Kernel32")]
        public static extern void AllocConsole();

        [DllImport("Kernel32", SetLastError = true)]
        public static extern void FreeConsole();

        public Parser()
        {
            this.LinesParsed = new List<String[]>();

            AllocConsole();
        }

        ~Parser() 
        {
            FreeConsole();
        }

        public void ParseFile()
        {
            StreamReader sr = new StreamReader("C:\\Users\\LaboTri-PC2\\Desktop\\dev\\test\\test.txt");
            Excel.Application excelApp = new Excel.Application();
            excelApp.Workbooks.Open("C:\\Users\\LaboTri-PC2\\Desktop\\dev\\test\\test.xlsx");

            try
            {
                String line;
                String[] words;
                int nbLine = 55;

                while ((line = sr.ReadLine()) != null)
                {
                    words = line.Split(' ');
                    words = (String[])words.Where((item, index) => item != "" && item != " ").ToArray();

                    for(int i = 0; i < words.Length; i++) 
                    {
                        Console.WriteLine(words[i]);
                    }

                    this.LinesParsed.Add(words);

                    nbLine++;
                }

                sr.Close();
                excelApp.Workbooks.Close();
            }
            catch
            {
                Console.WriteLine("Erreur : le fichier n'a pas pu être parsé correctement.");
                sr.Close();
                excelApp.Workbooks.Close();
            }

        }
    }
}
