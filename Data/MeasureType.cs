using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Data
{
    internal class MeasureType
    {
        public required String Name { get; set; }
        public int NominalValueIndex { get; set; }
        public int TolPlusIndex { get; set; }
        public int ValueIndex { get; set; }
        public int TolMinusIndex { get; set; }
        public required String Symbol { get; set; }

        public Data CreateData(List<double> values)
        {
            Data data = new Data();

            if (this.NominalValueIndex != -1) 
                data.SetNominalValue(values[this.NominalValueIndex]);

            if (this.TolPlusIndex != -1)
                data.SetTolPlus(values[this.TolPlusIndex]);

            if (this.ValueIndex != -1)
                data.SetValue(values[this.ValueIndex]);

            if (this.TolMinusIndex != -1)
                data.SetTolMinus(values[this.TolMinusIndex]);

            data.SetSymbol(this.Symbol);

            return data;
        }

        public String GetName()
        {
            return this.Name;
        }
    }
}
