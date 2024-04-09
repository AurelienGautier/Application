using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Application.Data
{
    internal class DataPosition : Data
    {
        public DataPosition() : base()
        {

        }

        public override double getNominalValue()
        {
            return 0.0;
        }

        public override double getValue() 
        {
            return base.GetValues()[3];
        }

        public override double getTolMinus()
        {
            return 0.0;
        }
    }
}
