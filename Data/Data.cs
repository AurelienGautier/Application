using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Data
{
    internal abstract class Data
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

        public abstract void WriteData(Excel.Application excelApp);

        public void PrintValues()
        {
            foreach (Double value in this.values)
            {
                Console.WriteLine(value);
            }
        }
    }
}
