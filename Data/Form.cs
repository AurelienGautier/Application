namespace Application.Data
{
    enum FormType
    {
        OnePiece,
        FivePieces,
        Capability
    }

    enum DataFrom
    {
        File,
        Files,
        Folder
    }

    class Form
    {
        public String Name { get; set; }
        public String Path { get; set; }
        public int DesignLine { get; set; }
        public int FirstLine { get; set; }
        public int FirstColumn { get; set; }
        public int LineToSign { get; set; }
        public int ColumnToSign { get; set; }
        public bool Modify { get; set; }
        public bool Sign { get; set; }
        public FormType Type { get; set; }
        public DataFrom DataFrom { get; set; }
        public int ClientLine { get; set; }
        public int StandardLine { get; set; }
        public List<int>? CapabilityMeasureNumber { get; set; }

        public Form(String name, String path, int designLine, int firstLine, int firstColumn, int lineToSign, int columnToSign, FormType type, DataFrom dataFrom, int clientLine, int standardLine)
        {
            this.Name = name;
            this.Path = path;
            this.DesignLine = designLine;
            this.FirstLine = firstLine;
            this.FirstColumn = firstColumn;
            this.LineToSign = lineToSign;
            this.ColumnToSign = columnToSign;
            this.Modify = false;
            this.Sign = false;
            Type = type;
            DataFrom = dataFrom;
            ClientLine = clientLine;
            StandardLine = standardLine;
        }
    }
}
