using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application
{
    enum MEASURE_INFO
    {
        NOMINAL_VALUE,
        TOL_PLUS,
        VALUE,
        TOL_MINUS
    }

    internal class ConfigSingleton
    {
        private static ConfigSingleton instance = null;
        private readonly List<String> measureTypesNames;
        private readonly List<List<int>> measureTypesValues;
        private readonly List<String> measureTypesSymbols;

        private ConfigSingleton()
        {
            this.measureTypesNames = new List<string>();
            this.measureTypesValues = new List<List<int>>();
            this.measureTypesSymbols = new List<string>();

            this.getMeasureDataFromFile();
        }

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

        private void getMeasureDataFromFile()
        {
            // Valeurs banales
            this.addDataType("Distance", new List<int> { 0, 1, 2, 4 }, "");
            this.addDataType("Diameter", new List<int> { 0, 1, 2, 4 }, "⌀");
            this.addDataType("Pos.X", new List<int> { 0, 1, 2, 4 }, "X");
            this.addDataType("Pos.Y", new List<int> { 0, 1, 2, 4 }, "Y");
            this.addDataType("Pos.Z", new List<int> { 0, 1, 2, 4 }, "Z");
            this.addDataType("Angle", new List<int> { 0, 1, 2, 4 }, "");
            this.addDataType("Result", new List<int> { 0, 1, 2, 4 }, "");
            this.addDataType("Min.Ax/2", new List<int> { 0, 1, 2, 4 }, "");
            this.addDataType("X-Angle", new List<int> { 0, 1, 2, 4 }, "");

            // Valeurs spéciales
            this.addDataType("Ax:R/Out", new List<int> { 0, 1, 2, -1 }, "");
            this.addDataType("CirR/Out", new List<int> { 0, 1, 2, -1 }, "");
            this.addDataType("Symmetry", new List<int> { 0, 1, 2, -1 }, "");

            this.addDataType("Cylinder", new List<int> { 0, 0, 1, -1 }, "");

            this.addDataType("Concentr", new List<int> { -1, 1, 3, -1 }, "");
            this.addDataType("Position", new List<int> { -1, 1, 3, -1 }, "⊕");

            this.addDataType("Flatness", new List<int> { -1, 0, 1, -1 }, "⏥");
            this.addDataType("Rectang.", new List<int> { -1, 0, 1, -1 }, "_");
            this.addDataType("Parallele", new List<int> { -1, 0, 1, -1 }, "//");
        }

        private void addDataType(String type, List<int> values, String symbol)
        {
            this.measureTypesNames.Add(type);
            this.measureTypesValues.Add(values);
            this.measureTypesSymbols.Add(symbol);
        }

        private int getDataIndex(String type)
        {
            for(int i = 0; i < this.measureTypesNames.Count; i++)
            {
                if (this.measureTypesNames[i] == type) return i;
            }

            return -1;
        }

        private Data.Data createData(int index, List<double> values)
        {
            Data.Data data = new Data.Data();

            int nominalValueIndex = this.measureTypesValues[index][(int)MEASURE_INFO.NOMINAL_VALUE];
            int tolPlusIndex = this.measureTypesValues[index][(int)MEASURE_INFO.TOL_PLUS];
            int valueIndex = this.measureTypesValues[index][(int)MEASURE_INFO.VALUE];
            int tolMinusIndex = this.measureTypesValues[index][(int)MEASURE_INFO.TOL_MINUS];

            nominalValueIndex = nominalValueIndex == -1 ? 0 : nominalValueIndex;
            tolPlusIndex = tolPlusIndex == -1 ? 0 : tolPlusIndex;
            valueIndex = valueIndex == -1 ? 0 : valueIndex;
            tolMinusIndex = tolMinusIndex == -1 ? 0 : tolMinusIndex;

            data.SetNominalValue(values[nominalValueIndex]);
            data.SetTolPlus(values[tolPlusIndex]);
            data.SetValue(values[valueIndex]);
            data.SetTolMinus(values[tolMinusIndex]);
            data.SetSymbol(this.measureTypesSymbols[index]);

            return data;
        }

        public Data.Data GetData(List<String> line, List<double> values)
        {
            if (line[2] == "Pos.") line[2] += line[3];

            int index = this.getDataIndex(line[2]);

            if(index == -1) throw new Exceptions.MeasureTypeNotFoundException();

            Data.Data data = this.createData(index, values);

            return data;
        }
    }
}
