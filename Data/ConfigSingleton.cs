using Application.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using static System.Net.Mime.MediaTypeNames;
using System.IO;
using System.Data;

namespace Application.Data
{
    internal class ConfigSingleton
    {
        private static ConfigSingleton? instance = null;
        private readonly List<MeasureType> measureTypes;

        public String Signature { get; set; }

        /*-------------------------------------------------------------------------*/

        private ConfigSingleton()
        {
            this.measureTypes = new List<MeasureType>();

            this.Signature = this.getSignature();

            this.getMeasureDataFromFile();
        }

        /*-------------------------------------------------------------------------*/

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

        /*-------------------------------------------------------------------------*/

        private String getSignature()
        {
            String json = this.getFileContent(Environment.CurrentDirectory + "\\conf\\conf.json");

            Dictionary<String, String>? data = JsonConvert.DeserializeObject<Dictionary<String, String>>(json);

            if (data == null)
                throw new Exceptions.ConfigDataException("Une erreur s'est produite lors de la récupération de la signature.");

            return data["Signature"];
        }

        /*-------------------------------------------------------------------------*/

        private void getMeasureDataFromFile()
        {
            String filePath = Environment.CurrentDirectory + "\\conf\\measureTypes.json";

            String json = this.getFileContent(filePath);

            DataSet? dataSet = JsonConvert.DeserializeObject<DataSet>(json);

            if (dataSet == null) 
                throw new Exceptions.ConfigDataException("Une erreur s'est produite lors de la récupération des types de mesure.");

            DataTable? dataTable = dataSet.Tables["Measures"];

            if (dataTable == null) 
                throw new Exceptions.ConfigDataException("La syntaxe du fichier contenant les types de mesure est incorrecte.");

            foreach (DataRow row in dataTable.Rows)
            {
                this.addData(row);
            }
        }

        /*-------------------------------------------------------------------------*/

        private void addData(DataRow row)
        {
            String? name = row["Name"].ToString();
            String? nominalValueIndex = row["NominalValueIndex"].ToString();
            String? tolPlusIndex = row["TolPlusIndex"].ToString();
            String? valueIndex = row["ValueIndex"].ToString();
            String? tolMinusIndex = row["TolMinusIndex"].ToString();
            String? symbol = row["Symbol"].ToString();

            if(name == null || nominalValueIndex == null || tolPlusIndex == null || valueIndex == null || tolMinusIndex == null || symbol == null)
                throw new Exceptions.ConfigDataException("Il existe au moins un type de mesure dont la syntaxe n'est pas correcte. Veuillez vérifier le contenu du fichier de configuration.");

            this.measureTypes.Add(new MeasureType
            {
                Name = name,
                NominalValueIndex = int.Parse(nominalValueIndex),
                TolPlusIndex = int.Parse(tolPlusIndex),
                ValueIndex = int.Parse(valueIndex),
                TolMinusIndex = int.Parse(tolMinusIndex),
                Symbol = symbol
            });
        }

        /*-------------------------------------------------------------------------*/

        private String getFileContent(String filePath)
        {
            StreamReader reader = new StreamReader(filePath);
            String content = reader.ReadToEnd();
            reader.Close();

            return content;
        }

        /*-------------------------------------------------------------------------*/

        public MeasureType? GetMeasureTypeFromLibelle(String libelle)
        {
            foreach (MeasureType measureType in this.measureTypes)
            {
                if (measureType.GetName() == libelle) return measureType;
            }

            return null;
        }

        /*-------------------------------------------------------------------------*/

        public Data? GetData(List<String> line, List<double> values)
        {
            if (line[2] == "Pos.") line[2] += line[3];

            MeasureType? measureType = this.GetMeasureTypeFromLibelle(line[2]);

            if (measureType == null)
                return null;

            return measureType.CreateData(values);
        }

        /*-------------------------------------------------------------------------*/

        public List<MeasureType> GetMeasureTypes()
        {
            return this.measureTypes;
        }

        /*-------------------------------------------------------------------------*/

        public void SetSignature(String signature)
        {
            this.Signature = signature;

            Dictionary<String, String> data = new Dictionary<String, String>
            {
                { "Signature", signature }
            };

            String json = JsonConvert.SerializeObject(data);

            StreamWriter writer = new StreamWriter(Environment.CurrentDirectory + "\\conf\\conf.json");
            writer.Write(json);
            writer.Close();
        }

        /*-------------------------------------------------------------------------*/
    }
}
