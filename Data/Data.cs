using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Data
{
    internal class Data
    {
        List<Double> values;

        public Data() 
        {
            this.values = new List<Double>();
        }

        public void setValues(List<Double> values)
        {
            this.values.AddRange(values);
        }

        public void WriteData(Excel.Application excelApp)
        {

        }

        public double getNominalValue()
        {
            return this.values[0];
        }

        public double getTolPlus()
        {
            return this.values[1];
        }

        public double getValue()
        {
            return this.values[2];
        }

        public double getEcart()
        {
            return this.values[3];
        }

        public double getTolMinus() 
        {
            return this.values[4];
        }

        public void PrintValues()
        {
            foreach (Double value in this.values)
            {
                Console.WriteLine(value);
            }
        }
    }
}
