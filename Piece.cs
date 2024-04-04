using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
