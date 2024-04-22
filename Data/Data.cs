namespace Application.Data
{
    internal class Data
    {
        private double nominalValue;
        private double tolerancePlus;
        private double toleranceMinus;
        private double value;

        // Le symbole du type de mesure
        private String symbol;

        public Data()
        {
            this.nominalValue = 0.0;
            this.tolerancePlus = 0.0;
            this.toleranceMinus = 0.0;
            this.value = 0.0;
            this.symbol = "";
        }

        public void SetNominalValue(double nominalValue)
        {
            this.nominalValue = nominalValue;
        }

        public double GetNominalValue()
        {
            return this.nominalValue;
        }

        public void SetTolPlus(double tolerancePlus)
        {
            this.tolerancePlus = tolerancePlus;
        }

        public double GetTolPlus()
        {
            return this.tolerancePlus;
        }

        public void SetValue(double value)
        {
            this.value = value;
        }

        public double GetValue()
        {
            return this.value;
        }

        public void SetTolMinus(double toleranceMinus)
        {
            this.toleranceMinus = toleranceMinus;
        }

        public double GetTolMinus() 
        {
            return this.toleranceMinus;
        }

        public void SetSymbol(String symbol)
        {
            this.symbol = symbol;
        }

        public String GetSymbol()
        {
            return this.symbol;
        }
    }
}
