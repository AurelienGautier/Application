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

        /**
         * ConfigSingleton
         * 
         * Constructeur de la classe (privé car singleton donc doit être inaccessible de l'extérieur de la classe)
         * 
         */
        private ConfigSingleton()
        {
            this.measureTypes = new List<MeasureType>();

            this.Signature = this.getSignature();

            this.getMeasureDataFromFile();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Singleton instance
         * 
         * Retourne l'instance du singleton et la crée si elle n'existe pas
         * 
         */
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

        /**
         * getSignature
         * 
         * Récupère la signature dans le fichier de configuration
         * 
         * @return String
         * 
         */
        private String getSignature()
        {
            String json = this.getFileContent(Environment.CurrentDirectory + "\\conf\\conf.json");

            Dictionary<String, String>? data = JsonConvert.DeserializeObject<Dictionary<String, String>>(json);

            if (data == null)
                throw new Exceptions.ConfigDataException("Une erreur s'est produite lors de la récupération de la signature.");

            return data["Signature"];
        }

        /*-------------------------------------------------------------------------*/

        /**
         * getMeasureDataFromFile
         * 
         * Récupère les types de mesure depuis le fichier de configuration
         * 
         */
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

        /**
         * addData
         * 
         * Désérialise une ligne du fichier de configuration et l'ajoute dans la liste des types de mesure
         * 
         * @param DataRow row
         * 
         */
        private void addData(DataRow row)
        {
            // Récupération de chaque champ de la ligne
            String? name = row["Name"].ToString();
            String? nominalValueIndex = row["NominalValueIndex"].ToString();
            String? tolPlusIndex = row["TolPlusIndex"].ToString();
            String? valueIndex = row["ValueIndex"].ToString();
            String? tolMinusIndex = row["TolMinusIndex"].ToString();
            String? symbol = row["Symbol"].ToString();

            // Lance une exception si un champ est null
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

        /**
         * getFileContent
         * 
         * Récupère le contenu entier d'un fichier
         * 
         * @param String filePath - Chemin du fichier dont le contenu est à récupérer
         * 
         * @return String
         * 
         */
        private String getFileContent(String filePath)
        {
            StreamReader reader = new StreamReader(filePath);
            String content = reader.ReadToEnd();
            reader.Close();

            return content;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * GetMeasureTypeFromLibelle
         * 
         * Retourne le type de mesure dont le libellé est passé en paramètre
         * 
         * @param String libelle
         * 
         * @return MeasureType? - Le type de mesure ou null s'il n'existe pas
         * 
         */
        public MeasureType? GetMeasureTypeFromLibelle(String libelle)
        {
            foreach (MeasureType measureType in this.measureTypes)
            {
                if (measureType.Name == libelle) return measureType;
            }

            return null;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * GetData
         * 
         * Crée un objet Data à partir des valeurs passées en paramètre
         * 
         * @param List<String> line - Ligne du fichier de données
         * @param List<double> values - Toutes les valeurs relatives à la ligne
         * 
         * @return Data? - L'objet Data créé ou null si le type de mesure n'existe pas
         * 
         */
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

        /**
         * SetSignature
         * 
         * Modifie la signature dans le fichier de configuration et dans l'instance
         * 
         * @param String signature
         * 
         */
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

        /**
         * UpdateMeasureType
         * 
         * Modifie un type de mesure dans le fichier de configuration et dans l'instance
         * 
         * @param MeasureType measureType - Le type de mesure à modifier
         * 
         */
        public void UpdateMeasureType(MeasureType measureType, MeasureType newMeasureType)
        {
            for (int i = 0; i < this.measureTypes.Count; i++)
            {
                if (this.measureTypes[i] == measureType)
                {
                    this.measureTypes[i] = newMeasureType;
                    break;
                }
            }

            this.serializeMeasureTypes();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * DeleteMeasureType
         * 
         * Supprime un type de mesure dans le fichier de configuration et dans l'instance
         * 
         * @param MeasureType measureType - Le type de mesure à supprimer
         * 
         */
        public void DeleteMeasureType(String libelle)
        {
            for (int i = 0; i < this.measureTypes.Count; i++)
            {
                if (this.measureTypes[i].Name == libelle)
                {
                    this.measureTypes.RemoveAt(i);
                    break;
                }
            }

            this.serializeMeasureTypes();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * serializeMeasureTypes
         * 
         * Convertir les types de mesure en JSON et les écrit dans le fichier de configuration
         * 
         */
        private void serializeMeasureTypes()
        {
            DataSet dataSet = new DataSet();
            DataTable dataTable = new DataTable("Measures");
            dataTable.Columns.Add("Name");
            dataTable.Columns.Add("NominalValueIndex");
            dataTable.Columns.Add("TolPlusIndex");
            dataTable.Columns.Add("ValueIndex");
            dataTable.Columns.Add("TolMinusIndex");
            dataTable.Columns.Add("Symbol");

            dataSet.Tables.Add(dataTable);

            foreach (MeasureType measure in this.measureTypes)
            {
                DataRow row = dataTable.NewRow();
                row["Name"] = measure.Name;
                row["NominalValueIndex"] = measure.NominalValueIndex;
                row["TolPlusIndex"] = measure.TolPlusIndex;
                row["ValueIndex"] = measure.ValueIndex;
                row["TolMinusIndex"] = measure.TolMinusIndex;
                row["Symbol"] = measure.Symbol;

                dataTable.Rows.Add(row);
            }

            dataSet.AcceptChanges();

            String json = JsonConvert.SerializeObject(dataSet, Formatting.Indented);

            StreamWriter writer = new StreamWriter(Environment.CurrentDirectory + "\\conf\\measureTypes.json");
            writer.Write(json);
            writer.Close();
        }

        /*-------------------------------------------------------------------------*/
    }
}
