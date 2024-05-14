﻿using Newtonsoft.Json;
using System.IO;
using System.Data;
using System.Drawing;

namespace Application.Data
{
    internal class ConfigSingleton
    {
        private static ConfigSingleton? instance = null;
        private readonly List<MeasureType> measureTypes;
        public Image? Signature { get; set; }

        /*-------------------------------------------------------------------------*/

        /**
         * Constructeur de la classe (privé car singleton donc doit être inaccessible de l'extérieur de la classe)
         * 
         */
        private ConfigSingleton()
        {
            this.measureTypes = new List<MeasureType>();

            this.Signature = this.getSignatureFromFile();

            this.getMeasureDataFromFile();
        }

        /*-------------------------------------------------------------------------*/

        /**
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
         * Récupère la signature dans le fichier de configuration
         * 
         * @return Image? - La signature ou null si elle n'est pas correcte
         * 
         */
        private Image? getSignatureFromFile()
        {
            String json = this.getFileContent(Environment.CurrentDirectory + "\\conf\\conf.json");

            Dictionary<String, String>? data = JsonConvert.DeserializeObject<Dictionary<String, String>>(json);

            if (data == null)
                throw new Exceptions.ConfigDataException("Le fichier de configuration est incorrect ou a été déplacé.");
            else if(!data.ContainsKey("Signature"))
                throw new Exceptions.ConfigDataException("Le fichier de configuration ne contient pas le champ signature.");
            
            Image? signature;
            try
            {
                signature = Image.FromFile(data["Signature"]);
            }
            catch
            {
                signature = null;
            }

            return signature;
        }

        /*-------------------------------------------------------------------------*/

        /**
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

        /**
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
         * @param String signaturePath - Chemin vers la nouvelle signature
         * 
         */
        public void SetSignature(String signaturePath)
        {
            try
            {
                this.Signature = Image.FromFile(signaturePath);
            }
            catch
            {
                throw new Exceptions.ConfigDataException("Le chemin vers la signature est incorrect.");
            }

            Dictionary<String, String> data = new Dictionary<String, String>
            {
                { "Signature", signaturePath }
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
        public void UpdateMeasureType(MeasureType? measureType, MeasureType newMeasureType)
        {
            if(measureType == null)
            {
                this.measureTypes.Add(newMeasureType);
            }
            else
            {
                for (int i = 0; i < this.measureTypes.Count; i++)
                {
                    if (this.measureTypes[i] == measureType)
                    {
                        this.measureTypes[i] = newMeasureType;
                        break;
                    }
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

        public List<Form> GetMitutoyoForms()
        {
            List<Form> forms = new List<Form>();

            forms.Add(new Form("Rapport 1 pièce", Environment.CurrentDirectory + "\\form\\rapport1piece", 26, 30, 1, 55, 14, FormType.OnePiece, DataFrom.File, 18));
            forms.Add(new Form("Outillage de contrôle", Environment.CurrentDirectory + "\\form\\outillageDeControle", 25, 26, 1, 51, 14, FormType.OnePiece, DataFrom.File, 17));
            forms.Add(new Form("Rapport 5 pièces", Environment.CurrentDirectory + "\\form\\rapport5pieces", 25, 17, 1, 51, 14, FormType.FivePieces, DataFrom.Folder, 17));

            return forms;
        }

        /*-------------------------------------------------------------------------*/

        public List<Form> GetAyonisForms()
        {
            List<Form> forms = new List<Form>();

            forms.Add(new Form("Rapport 1 pièce", Environment.CurrentDirectory + "\\form\\rapport1piece", 26, 30, 1, 55, 14, FormType.OnePiece, DataFrom.File, 18));
            forms.Add(new Form("Rapport 5 pièces", Environment.CurrentDirectory + "\\form\\rapport5pieces", 25, 17, 1, 51, 14, FormType.FivePieces, DataFrom.File, 17));

            return forms;
        }

        /*-------------------------------------------------------------------------*/
    }
}
