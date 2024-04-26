namespace Application.Data
{
    internal class Data
    {
        public double NominalValue { get; set; }
        public double TolerancePlus { get; set; }
        public double ToleranceMinus { get; set; }
        public double Value { get; set; }
        public String Symbol { get; set; }

        public Data()
        {
            this.NominalValue = 0.0;
            this.TolerancePlus = 0.0;
            this.ToleranceMinus = 0.0;
            this.Value = 0.0;
            this.Symbol = "";
        }
    }
}
