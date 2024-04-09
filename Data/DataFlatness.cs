using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Data
{
    internal class DataFlatness : Data
    {
        public DataFlatness() : base()
        {
            
        }

        public override double getValue()
        {
            return base.GetValues()[1];
        }

        public override double getTolPlus()
        {
            return base.GetValues()[1];
        }

        public override double getTolMinus()
        {
            return 0.0;
        }
    }
}
