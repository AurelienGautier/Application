using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Data
{
    internal class DataAxCirOut : Data
    {
        public DataAxCirOut() : base()
        {

        }

        public override double GetValue()
        {
            return base.GetValues()[1];
        }

        public override double GetTolPlus()
        {
            return base.GetValues()[2];
        }

        public override double GetTolMinus()
        {
            return 0.0;
        }

        public override double GetOutTolerance()
        {
            if(base.GetValues().Count > 3) return base.GetValues()[3];

            return 0.0;
        }
    }
}
