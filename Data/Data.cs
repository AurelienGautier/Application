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

        public virtual double getNominalValue()
        {
            return this.values[0];
        }

        public virtual double getTolPlus()
        {
            return this.values[1];
        }

        public virtual double getValue()
        {
            return this.values[2];
        }

        public virtual double getEcart()
        {
            return this.values[3];
        }

        public virtual double getTolMinus() 
        {
            this.PrintValues();

            return this.values[4];
        }

        public virtual double getOutTolerance()
        {
            if(this.values.Count > 5)
            {
                return this.values[5];
            }

            return 0.0;
        }

        protected List<double> GetValues()
        {
            return this.values;
        }

        public void PrintValues()
        {
            foreach (Double value in this.values)
            {
                Console.WriteLine(value);
            }

            Console.WriteLine();
        }
    }
}
