using Newtonsoft.Json;
using System.IO;
using System.Data;
using System.Drawing;
using Application.Writers;

namespace Application.Data
{
    /// <summary>
    /// Represents a singleton configuration class.
    /// </summary>
    internal class ConfigSingleton
    {
        private static ConfigSingleton? instance = null;
        private readonly List<MeasureType> measureTypes;
        public Image? Signature { get; set; }
        private List<Standard> standards;
        private Dictionary<string, string> headerFieldsMatch;
        private Dictionary<string, string> pageNames;

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Private constructor for the ConfigSingleton class.
        /// </summary>
        private ConfigSingleton()
        {
            this.measureTypes = new List<MeasureType>();

            this.Signature = this.GetSignatureFromFile();

            this.standards = new List<Standard>();

            this.GetMeasureDataFromFile();

            this.GetStandardsFromJsonFile();

            this.headerFieldsMatch = new Dictionary<string, string>();

            this.getHeaderFieldsFromFile();

            this.pageNames = new Dictionary<string, string>();

            this.getPageNamesFromFile();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the instance of the ConfigSingleton class.
        /// </summary>
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

        /// <summary>
        /// Gets the signature from the configuration file.
        /// </summary>
        /// <returns>The signature image or null if it is not valid.</returns>
        private Image? GetSignatureFromFile()
        {
            string? json = this.GetFileContent(Environment.CurrentDirectory + "\\conf\\conf.json");

            if (json == null) return null;

            Dictionary<string, string>? data = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);

            if (data == null) return null;

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

        /// <summary>
        /// Gets the measure data from the configuration file.
        /// </summary>
        private void GetMeasureDataFromFile()
        {
            string filePath = Environment.CurrentDirectory + "\\conf\\measureTypes.json";

            string? json = this.GetFileContent(filePath);

            if (json == null) return;

            DataSet? dataSet = JsonConvert.DeserializeObject<DataSet>(json);

            if (dataSet == null)
                throw new Exceptions.ConfigDataException("Une erreur s'est produite lors de la récupération des types de mesure.");

            DataTable? dataTable = dataSet.Tables["Measures"];

            if (dataTable == null)
                throw new Exceptions.ConfigDataException("La syntaxe du fichier contenant les types de mesure est incorrecte.");

            foreach (DataRow row in dataTable.Rows)
            {
                this.AddData(row);
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Deserializes a row from the configuration file and adds it to the list of measure types.
        /// </summary>
        /// <param name="row">The row to add.</param>
        private void AddData(DataRow row)
        {
            // Get each field from the row
            string? name = row["Name"].ToString();
            string? nominalValueIndex = row["NominalValueIndex"].ToString();
            string? tolPlusIndex = row["TolPlusIndex"].ToString();
            string? valueIndex = row["ValueIndex"].ToString();
            string? tolMinusIndex = row["TolMinusIndex"].ToString();
            string? symbol = row["Symbol"].ToString();

            // Throw an exception if any field is null
            if (name == null || nominalValueIndex == null || tolPlusIndex == null || valueIndex == null || tolMinusIndex == null || symbol == null)
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

        /// <summary>
        /// Gets the content of a file.
        /// </summary>
        /// <param name="filePath">The path of the file to read.</param>
        /// <returns>The content of the file or null if it cannot be read.</returns>
        private string? GetFileContent(string filePath)
        {
            String content = "";

            try
            {
                StreamReader reader = new StreamReader(filePath);
                content = reader.ReadToEnd();
                reader.Close();
            }
            catch
            {
                return null;
            }

            return content;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Converts the measure types to JSON and writes them to the configuration file.
        /// </summary>
        private void SerializeMeasureTypes()
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

            this.writeJsonFile(Environment.CurrentDirectory + "\\conf\\measureTypes.json", dataSet);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Returns the measure type with the specified label.
        /// </summary>
        /// <param name="libelle">The label of the measure type.</param>
        /// <returns>The measure type or null if it does not exist.</returns>
        public MeasureType? GetMeasureTypeFromLibelle(string libelle)
        {
            foreach (MeasureType measureType in this.measureTypes)
            {
                if (measureType.Name == libelle) return measureType;
            }

            return null;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Creates a Measure object from the values passed as parameters.
        /// </summary>
        /// <param name="line">The line from the data file.</param>
        /// <param name="values">All the values related to the line.</param>
        /// <returns>The created Measure object or null if the measure type does not exist.</returns>
        public Measure? GetData(List<string> line, List<double> values)
        {
            if (line[2] == "Pos.") line[2] += line[3];

            MeasureType? measureType = this.GetMeasureTypeFromLibelle(line[2]);

            if (measureType == null)
                return null;

            return new Measure(measureType, values);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the list of measure types.
        /// </summary>
        /// <returns>The list of measure types.</returns>
        public List<MeasureType> GetMeasureTypes()
        {
            return this.measureTypes;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Sets the signature in the configuration file and in the instance.
        /// </summary>
        /// <param name="signaturePath">The path to the new signature.</param>
        public void SetSignature(string signaturePath)
        {
            try
            {
                this.Signature = Image.FromFile(signaturePath);
            }
            catch
            {
                throw new Exceptions.ConfigDataException("Le chemin vers la signature est incorrect.");
            }

            Dictionary<string, string> data = new Dictionary<string, string>
                {
                    { "Signature", signaturePath }
                };

            this.writeJsonFile(Environment.CurrentDirectory + "\\conf\\conf.json", data);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Updates a measure type in the configuration file and in the instance.
        /// </summary>
        /// <param name="measureType">The measure type to update.</param>
        /// <param name="newMeasureType">The new measure type.</param>
        public void UpdateMeasureType(MeasureType? measureType, MeasureType newMeasureType)
        {
            if (measureType == null)
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

            this.SerializeMeasureTypes();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Deletes a measure type from the configuration file and from the instance.
        /// </summary>
        /// <param name="libelle">The label of the measure type to delete.</param>
        public void DeleteMeasureType(string libelle)
        {
            for (int i = 0; i < this.measureTypes.Count; i++)
            {
                if (this.measureTypes[i].Name == libelle)
                {
                    this.measureTypes.RemoveAt(i);
                    break;
                }
            }

            this.SerializeMeasureTypes();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the list of Mitutoyo forms.
        /// </summary>
        /// <returns>The list of Mitutoyo forms.</returns>
        public List<Form> GetMitutoyoForms()
        {
            List<Form> forms = new List<Form>();

            forms.Add(new Form("Rapport 1 pièce", Environment.CurrentDirectory + "\\form\\rapport1piece", 26, 30, 1, 53, 11, FormType.OnePiece, DataFrom.File, 18, 75));
            forms.Add(new Form("Outillage de contrôle", Environment.CurrentDirectory + "\\form\\outillageDeControle", 26, 26, 1, 53, 11, FormType.OnePiece, DataFrom.File, 18, 75));
            forms.Add(new Form("Rapport 5 pièces", Environment.CurrentDirectory + "\\form\\rapport5pieces", 26, 17, 1, 51, 14, FormType.FivePieces, DataFrom.Files, 18, 75));
            forms.Add(new Form("Capabilité", Environment.CurrentDirectory + "\\form\\capabilite", 26, 103, 5, 53, 11, FormType.Capability, DataFrom.File, 18, 75));

            return forms;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the list of Ayonis forms.
        /// </summary>
        /// <returns>The list of Ayonis forms.</returns>
        public List<Form> GetAyonisForms()
        {
            List<Form> forms = new List<Form>();

            forms.Add(new Form("Rapport 1 pièce", Environment.CurrentDirectory + "\\form\\rapport1piece", 26, 30, 1, 53, 11, FormType.OnePiece, DataFrom.File, 18, 75));
            forms.Add(new Form("Rapport 5 pièces", Environment.CurrentDirectory + "\\form\\rapport5pieces", 26, 17, 1, 51, 14, FormType.FivePieces, DataFrom.File, 18, 75));

            return forms;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the standards from the Excel file.
        /// </summary>
        private void GetStandardsFromExcelFile()
        {
            ExcelLibraryLinkSingleton excelApiLink = ExcelLibraryLinkSingleton.Instance;
            string workbookPath = Environment.CurrentDirectory + "\\res\\etalons.xlsm";
            excelApiLink.OpenWorkBook(workbookPath);

            excelApiLink.ChangeWorkSheet(workbookPath, "raccordements à jour ");

            int currentLine = 9;

            while (excelApiLink.ReadCell(workbookPath, currentLine, 1) != "")
            {
                if (excelApiLink.MergedCells(workbookPath, currentLine, 1, currentLine, 2))
                {
                    currentLine++;
                }
                else
                {
                    string code = excelApiLink.ReadCell(workbookPath, currentLine, 1);
                    string name = excelApiLink.ReadCell(workbookPath, currentLine, 2);
                    string raccordement = excelApiLink.ReadCell(workbookPath, currentLine + 1, 2);
                    string validity = excelApiLink.ReadCell(workbookPath, currentLine + 2, 2).Substring(0, 10);
                    validity = validity.Substring(3);
                    validity = validity.Insert(0, "\'");

                    this.standards.Add(new Standard(code, name, raccordement, validity));

                    currentLine += 3;
                }
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the standards from the JSON file.
        /// </summary>
        private void GetStandardsFromJsonFile()
        {
            this.standards.Add(new Standard("", "", "", ""));

            String? json = this.GetFileContent(Environment.CurrentDirectory + "\\conf\\standards.json");

            if (json == null) return;

            List<Standard>? standardsFromFile = JsonConvert.DeserializeObject<List<Standard>>(json);

            if (standardsFromFile != null) this.standards = standardsFromFile;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Updates the standards in the configuration file.
        /// </summary>
        public void UpdateStandards()
        {
            this.standards.Clear();

            this.standards.Add(new Standard("", "", "", ""));

            this.GetStandardsFromExcelFile();

            this.writeJsonFile(Environment.CurrentDirectory + "\\conf\\standards.json", this.standards);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the list of standards.
        /// </summary>
        /// <returns>The list of the standards</returns>
        public List<Standard> GetStandards()
        {
            return this.standards;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the standard from the code specified in parameter.
        /// </summary>
        /// <param name="code">The code of the standard to get</param>
        /// <returns></returns>
        public Standard? GetStandardFromCode(String code)
        {
            if(this.standards != null)
                foreach (Standard standard in this.standards)
                {
                    if (standard.Code == code) return standard;
                }

            return null;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the header key name from the configuration file.
        /// </summary>
        /// <exception cref="Exceptions.ConfigDataException">If the header config data could not have been got from the config file</exception>
        private void getHeaderFieldsFromFile()
        {
            String? json = this.GetFileContent(Environment.CurrentDirectory + "\\conf\\headerFields.json");

            if (json == null) return;

            Dictionary<String, String>? headerFields = JsonConvert.DeserializeObject<Dictionary<String, String>>(json);

            if (headerFields == null)
                throw new Exceptions.ConfigDataException("Le fichier de configuration des champs d'en-tête est incorrect ou a été déplacé.");

            this.headerFieldsMatch = headerFields;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Returns the header key names
        /// </summary>
        /// <returns></returns>
        public Dictionary<String, String> GetHeaderFieldsMatch()
        {
            return this.headerFieldsMatch;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Sets the header value names
        /// </summary>
        /// <param name="designation"></param>
        /// <param name="planNb"></param>
        /// <param name="index"></param>
        /// <param name="clientName"></param>
        /// <param name="observationNum"></param>
        /// <param name="pieceReceptionDate"></param>
        /// <param name="observations"></param>
        public void SetHeaderFieldsMatch(String designation, String planNb, String index, String clientName, String observationNum, String pieceReceptionDate, String observations)
        {
            this.headerFieldsMatch["Designation"] = designation;
            this.headerFieldsMatch["PlanNb"] = planNb;
            this.headerFieldsMatch["Index"] = index;
            this.headerFieldsMatch["ClientName"] = clientName;
            this.headerFieldsMatch["ObservationNum"] = observationNum;
            this.headerFieldsMatch["PieceReceptionDate"] = pieceReceptionDate;
            this.headerFieldsMatch["Observations"] = observations;

            this.writeJsonFile(Environment.CurrentDirectory + "\\conf\\headerFields.json", this.headerFieldsMatch);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the page names from the configuration file to know which page is the header page and which one is the measure page.
        /// </summary>
        /// <exception cref="Exceptions.ConfigDataException"></exception>
        private void getPageNamesFromFile()
        {
            String? json = this.GetFileContent(Environment.CurrentDirectory + "\\conf\\pageNames.json");

            if (json == null) return;

            Dictionary<String, String>? pageNamesFromFile = JsonConvert.DeserializeObject<Dictionary<String, String>>(json);

            if (pageNamesFromFile == null)
                throw new Exceptions.ConfigDataException("Le fichier de configuration des noms de pages est incorrect ou a été déplacé.");

            this.pageNames = pageNamesFromFile;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Sets the page names in the configuration file.
        /// </summary>
        /// <param name="headerPage"></param>
        /// <param name="measurePage"></param>
        public void SetPageNames(String headerPage, String measurePage)
        {
            this.pageNames["HeaderPage"] = headerPage;
            this.pageNames["MeasurePage"] = measurePage;

            this.writeJsonFile(Environment.CurrentDirectory + "\\conf\\pageNames.json", this.pageNames);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Returns the page names
        /// </summary>
        /// <returns></returns>
        public Dictionary<String, String> GetPageNames()
        {
            return this.pageNames;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Replace the content of a json file with the data passed as parameter.
        /// </summary>
        /// <param name="filePath">The path of the file to modify</param>
        /// <param name="data">The data to write in the file</param>
        private void writeJsonFile(String filePath, Object data)
        {
            String json = JsonConvert.SerializeObject(data, Formatting.Indented);

            Directory.CreateDirectory(Environment.CurrentDirectory + "\\conf");
            File.Create(filePath).Close();
            StreamWriter writer = new StreamWriter(filePath);
            writer.Write(json);
            writer.Close();
        }

        /*-------------------------------------------------------------------------*/
    }
}
