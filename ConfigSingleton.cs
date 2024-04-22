using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application
{
    internal class ConfigSingleton
    {
        private static ConfigSingleton instance = null;

        public static ConfigSingleton Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new ConfigSingleton();
                }

                return instance;
            }
        }

        public Data.Data GetData(List<String> line, List<double> values)
        {
            Data.Data data = new Data.Data();

            if (line[2] == "Distance"
                || line[2] == "Diameter"
                || line[2] == "Pos."
                || line[2] == "Angle"
                || line[2] == "Result"
                || line[2] == "Min.Ax/2")
            {
                if (line[2] == "Diameter") data.SetSymbol("⌀");

                if (line[2] == "Pos.")
                {
                    data.SetSymbol(line[3]);
                    line[2] += line[3];
                    line.RemoveAt(3);
                }

                data.SetNominalValue(values[0]);
                data.SetTolPlus(values[1]);
                data.SetValue(values[2]);
                data.SetTolMinus(values[4]);
            }
            else if (line[2] == "Ax:R/Out" || line[2] == "CirR/Out" || line[2] == "Symmetry")
            {
                data.SetNominalValue(values[0]);
                data.SetTolPlus(values[2]);
                data.SetValue(values[1]);
            }
            else if (line[2] == "Concentr")
            {
                data.SetTolPlus(values[1]);
                data.SetValue(values[3]);

            }
            else if (line[2] == "Position")
            {
                data.SetSymbol("⊕");
                data.SetTolPlus(values[1]);
                data.SetValue(values[3]);
            }
            else if (line[2] == "Flatness" || line[2] == "Rectang." || line[2] == "Parallele")
            {
                if (line[2] == "Flatness") data.SetSymbol("⏥");
                else if (line[2] == "Rectang.") data.SetSymbol("_");
                else if (line[2] == "Parallel") data.SetSymbol("//");

                data.SetValue(values[1]);
                data.SetTolPlus(values[0]);
            }

            return data;
        }
    }
}
