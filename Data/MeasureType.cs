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
    }
}
