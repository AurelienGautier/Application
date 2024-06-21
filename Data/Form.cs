namespace Application.Data
{
    public enum FormType
    {
        OnePiece,
        FivePieces,
        Capability
    }

    public enum DataFrom
    {
        File,
        Files,
        Folder
    }

    public class Form(String name, String path, int designLine, int firstLine, int firstColumn, int lineToSign, int columnToSign, FormType type, DataFrom dataFrom, int clientLine, int standardLine, MeasureMachine measureMachine)
    {
        public String Name { get; set; } = name;
        public String Path { get; set; } = path;
        public int DesignLine { get; set; } = designLine;
        public int FirstLine { get; set; } = firstLine;
        public int FirstColumn { get; set; } = firstColumn;
        public int LineToSign { get; set; } = lineToSign;
        public int ColumnToSign { get; set; } = columnToSign;
        public bool Modify { get; set; } = false;
        public bool Sign { get; set; } = false;
        public FormType Type { get; set; } = type;
        public DataFrom DataFrom { get; set; } = dataFrom;
        public int ClientLine { get; set; } = clientLine;
        public int StandardLine { get; set; } = standardLine;
        public List<int> CapabilityMeasureNumber { get; set; } = [];
        public List<String> SourceFiles { get; set; } = [];
        public MeasureMachine MeasureMachine { get; set; } = measureMachine;
        public List<Standard> Standards { get; set; } = [];
        public string DestinationPath { get; set; } = "";

        public Form Copy()
        {
            return (Form)this.MemberwiseClone();
        }
    }
}
