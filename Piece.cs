using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application
{
    internal class Piece
    {
        private List<List<Data.Data>> pieceData;
        private List<String> measureTypes;

        public Piece() 
        {
            this.pieceData = new List<List<Data.Data>>();
            this.measureTypes = new List<String>();
        }

        public void AddMeasureType(String measureType)
        {
            this.measureTypes.Add(measureType);
            this.pieceData.Add(new List<Data.Data>());
        }

        public void AddData(Data.Data data)
        {
            this.pieceData[pieceData.Count - 1].Add(data);
        }

        public void setValues(List<double> values)
        {
            int i = pieceData.Count - 1;
            int j = this.pieceData[i].Count - 1;

            this.pieceData[i][j]
                .setValues(values);
        }

        public void WriteBaseValues(Excel.Application excelApp, char col, int line)
        {
            for(int i = 0; i < pieceData.Count; i++) 
            {
                /*excelApp.Range["A1"].Value = "coucou les amis";*/
                excelApp.Range[col + line.ToString()].Value = this.measureTypes[i];
                line++;

                for(int j = 0; j < pieceData[i].Count; j++)
                {
                    excelApp.Range[col + line.ToString()].Value = this.pieceData[i][j].getNominalValue();
                    col += (char)2;
                    excelApp.Range[col + line.ToString()].Value = this.pieceData[i][j].getTolPlus();
                    col++;
                    excelApp.Range[col + line.ToString()].Value = this.pieceData[i][j].getTolMinus();

                    line++;
                    col -= (char)3;
                }
            }
        }

        public void WriteValues(Excel.Application excelApp, char col, int line)
        {
            for(int i = 0; i < pieceData.Count; i++)
            {
                line++;

                for (int j = 0; j < this.pieceData[i].Count;j++)
                {
                    Console.WriteLine(col + line.ToString());
                    excelApp.Range[col + line.ToString()].Value = this.pieceData[i][j].getValue();
                    col++;
                    excelApp.Range[col + line.ToString()].Value = this.pieceData[i][j].getEcart();
                    line++;
                    col--;
                }
            }
        }

        public void PrintTrucs()
        {
            for(int i = 0; i < pieceData.Count; i++) 
            {
                Console.WriteLine(measureTypes[i]);

                foreach(Data.Data data in pieceData[i]) 
                {
                    if(data != null) data.PrintValues();
                    Console.WriteLine();
                }
            }
        }
    }
}
