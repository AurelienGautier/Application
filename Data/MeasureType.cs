using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Data
{
    public class MeasureType
    {
        public required String Name { get; set; }
        public int NominalValueIndex { get; set; }
        public int TolPlusIndex { get; set; }
        public int ValueIndex { get; set; }
        public int TolMinusIndex { get; set; }
        public required String Symbol { get; set; }

        /**
         * CreateData
         * 
         * Créer un objet Data à partir d'une liste de valeurs et des index en attribut
         * 
         */
        public Data CreateData(List<double> values)
        {
            Data data = new Data();

            if (this.NominalValueIndex != -1) 
                data.NominalValue = values[this.NominalValueIndex];

            if (this.TolPlusIndex != -1)
                data.TolerancePlus = values[this.TolPlusIndex];

            if (this.ValueIndex != -1)
                data.Value = values[this.ValueIndex];

            if (this.TolMinusIndex != -1)
                data.ToleranceMinus = values[this.TolMinusIndex];

            data.Symbol = this.Symbol;

            return data;
        }
    }
}
