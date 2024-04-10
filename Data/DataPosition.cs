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

        public override double GetNominalValue()
        {
            return 0.0;
        }

        public override double GetValue() 
        {
            return base.GetValues()[3];
        }

        public override double GetTolMinus()
        {
            return 0.0;
        }
    }
}
