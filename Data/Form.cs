namespace Application.Data
{
    class Form
    {
        public String Name { get; set; }
        public String Path { get; set; }
        public int DesignLine { get; set; }
        public int FirstLine { get; set; }
        public int LineToSign { get; set; }
        public int ColumnToSign { get; set; }
        public bool Modify { get; set; }
        public bool Sign { get; set; }

        public Form(String name, String path, int designLine, int firstLine, int lineToSign, int columnToSign)
        {
            this.Name = name;
            this.Path = path;
            this.DesignLine = designLine;
            this.FirstLine = firstLine;
            this.LineToSign = lineToSign;
            this.ColumnToSign = columnToSign;
            this.Modify = false;
            this.Sign = false;
        }
    }
}
