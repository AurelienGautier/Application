using System.Data;
using System.Drawing;
using Application.Facade;
using Application.Parser;

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
        readonly private List<Standard> standards;
        readonly private Dictionary<string, string> headerFieldsMatch;
        readonly private Dictionary<string, string> pageNames;
        public List<MeasureMachine> Machines { get; set; }

        readonly string standardsFilePath = Environment.CurrentDirectory + "\\conf\\standards.json";
        readonly string headerFieldsFilePath = Environment.CurrentDirectory + "\\conf\\headerFields.json";
        readonly string pageNamesFilePath = Environment.CurrentDirectory + "\\conf\\pageNames.json";

        /// <summary>
        /// Private constructor for the ConfigSingleton class.
        /// </summary>
        private ConfigSingleton()
        {
            this.measureTypes = [];

            this.Signature = GetSignatureFromFile();

            this.GetMeasureDataFromFile();

            this.standards = JsonLibraryLink.GetJsonFilecontent<List<Standard>>(standardsFilePath);

            // Add an empty standard at the beginning of the list
            this.standards.Insert(0, new Standard("", "", "", ""));

            this.headerFieldsMatch = JsonLibraryLink.GetJsonFilecontent<Dictionary<String, String>>(this.headerFieldsFilePath);

            this.pageNames = JsonLibraryLink.GetJsonFilecontent<Dictionary<String, String>>(pageNamesFilePath);

            this.Machines = [];

            this.GetMachinesAndForms();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the instance of the ConfigSingleton class.
        /// </summary>
        public static ConfigSingleton Instance
        {
            get
            {
                instance ??= new ConfigSingleton();

                return instance;
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the signature from the configuration file.
        /// </summary>
        /// <returns>The signature image or null if it is not valid.</returns>
        private static Image? GetSignatureFromFile()
        {
            string filePath = Environment.CurrentDirectory + "\\conf\\conf.json";

            var data = JsonLibraryLink.GetJsonFilecontent<Dictionary<String, String>>(filePath);

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

            DataSet dataSet = JsonLibraryLink.GetJsonFilecontent<DataSet>(filePath);

            DataTable? dataTable = dataSet.Tables["Measures"] ?? throw new Exceptions.ConfigDataException("La syntaxe du fichier contenant les types de mesure est incorrecte.");

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
        /// Converts the measure types to JSON and writes them to the configuration file.
        /// </summary>
        private void SerializeMeasureTypes()
        {
            DataSet dataSet = new();
            DataTable dataTable = new("Measures");
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

            JsonLibraryLink.WriteJsonFile(Environment.CurrentDirectory + "\\conf\\measureTypes.json", dataSet);
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

            Dictionary<string, string> data = new()
            {
                    { "Signature", signaturePath }
                };

            JsonLibraryLink.WriteJsonFile(Environment.CurrentDirectory + "\\conf\\conf.json", data);
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
        /// Sets the machines and the reports of the application
        /// </summary>
        private void GetMachinesAndForms()
        {
            MeasureMachine mitutoyoMachine = new ("Mitutoyo", new TextFileParser());
            List<Form> mitutoyoForms =
            [
                new Form("Rapport 1 pièce", Environment.CurrentDirectory + "\\form\\rapport1piece", 26, 30, 1, 53, 11, FormType.OnePiece, DataFrom.File, 18, 75, mitutoyoMachine),
                new Form("Outillage de contrôle", Environment.CurrentDirectory + "\\form\\outillageDeControle", 26, 26, 1, 53, 11, FormType.OnePiece, DataFrom.File, 18, 75, mitutoyoMachine),
                new Form("Rapport 5 pièces", Environment.CurrentDirectory + "\\form\\rapport5pieces", 26, 17, 1, 51, 14, FormType.FivePieces, DataFrom.Files, 18, 75, mitutoyoMachine),
                new Form("Capabilité", Environment.CurrentDirectory + "\\form\\capabilite", 26, 103, 5, 53, 11, FormType.Capability, DataFrom.File, 18, 75, mitutoyoMachine),
            ];
            mitutoyoMachine.setPossibleForms(mitutoyoForms);

            MeasureMachine ayonisMachine = new ("Ayonis", new ExcelParser());
            List<Form> ayonisForms =
            [
                new Form("Rapport 1 pièce", Environment.CurrentDirectory + "\\form\\rapport1piece", 26, 30, 1, 53, 11, FormType.OnePiece, DataFrom.File, 18, 75, ayonisMachine),
                new Form("Rapport 5 pièces", Environment.CurrentDirectory + "\\form\\rapport5pieces", 26, 17, 1, 51, 14, FormType.FivePieces, DataFrom.File, 18, 75, ayonisMachine),
            ];
            ayonisMachine.setPossibleForms(ayonisForms);

            this.Machines =
            [
                mitutoyoMachine,
                ayonisMachine
            ];
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Gets the standards from the Excel file.
        /// </summary>
        private void GetStandardsFromExcelFile()
        {
            this.standards.Add(new Standard("", "", "", ""));

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
                    string validity = excelApiLink.ReadCell(workbookPath, currentLine + 2, 2)[..10];
                    validity = validity[3..];
                    validity = validity.Insert(0, "\'");

                    this.standards.Add(new Standard(code, name, raccordement, validity));

                    currentLine += 3;
                }
            }

            excelApiLink.CloseWorkBook(workbookPath);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Updates the standards in the configuration file.
        /// </summary>
        public void UpdateStandards()
        {
            this.standards.Clear();

            this.GetStandardsFromExcelFile();

            JsonLibraryLink.WriteJsonFile(Environment.CurrentDirectory + "\\conf\\standards.json", this.standards);
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

            JsonLibraryLink.WriteJsonFile(Environment.CurrentDirectory + "\\conf\\headerFields.json", this.headerFieldsMatch);
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

            JsonLibraryLink.WriteJsonFile(Environment.CurrentDirectory + "\\conf\\pageNames.json", this.pageNames);
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
    }
}
