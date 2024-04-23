using Application.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application.Data
{
    internal class ConfigSingleton
    {
        private static ConfigSingleton instance = null;
        private readonly List<MeasureType> measureTypes;

        private ConfigSingleton()
        {
            this.measureTypes = new List<MeasureType>();

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

        private void addDataType(String type, List<int> indexes, String symbol)
        {
            this.measureTypes.Add(new MeasureType
            {
                Name = type,
                NominalValueIndex = indexes[0],
                TolPlusIndex = indexes[1],
                ValueIndex = indexes[2],
                TolMinusIndex = indexes[3],
                Symbol = symbol
            });
        }

        private int getDataIndex(String type)
        {
            for(int i = 0; i < this.measureTypes.Count; i++)
            {
                if (this.measureTypes[i].GetName() == type) return i;
            }

            return -1;
        }

        public Data GetData(List<String> line, List<double> values)
        {
            if (line[2] == "Pos.") line[2] += line[3];

            int index = this.getDataIndex(line[2]);

            if(index == -1) throw new Exceptions.MeasureTypeNotFoundException();

            Data data = this.measureTypes[index].CreateData(values);

            return data;
        }

        public List<MeasureType> GetMeasureTypes()
        {
            return this.measureTypes;
        }
    }
}
