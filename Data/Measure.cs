using Application.Exceptions;

namespace Application.Data
{
    public class Measure
    {
        public double NominalValue { get; set; }
        public double TolerancePlus { get; set; }
        public double ToleranceMinus { get; set; }
        public double Value { get; set; }
        public MeasureType MeasureType { get; set; }

        public Measure(MeasureType measureType, List<double> values)
        {
            try
            {
                this.NominalValue = measureType.NominalValueIndex != -1 ? values[measureType.NominalValueIndex] : 0.0;
            }
            catch (System.ArgumentOutOfRangeException)
            {
                throw new ConfigDataException("L'indice de la valeur nominale du type de mesure " + measureType.Name + " n'est pas correcte, veuillez la modifier.");
            }

            try
            {
                this.TolerancePlus = measureType.TolPlusIndex != -1 ? values[measureType.TolPlusIndex] : 0.0;
            }
            catch (System.ArgumentOutOfRangeException)
            {
                throw new ConfigDataException("L'indice de la tolérance supérieure du type de mesure " + measureType.Name + " n'est pas correcte, veuillez la modifier.");
            }

            try
            {
                this.ToleranceMinus = measureType.TolMinusIndex != -1 ? values[measureType.TolMinusIndex] : 0.0;
            }
            catch (System.ArgumentOutOfRangeException)
            {
                throw new ConfigDataException("L'indice de la tolérance inférieure du type de mesure " + measureType.Name + " n'est pas correcte, veuillez la modifier.");
            }

            try
            {
                this.Value = measureType.ValueIndex != -1 ? values[measureType.ValueIndex] : 0.0;
            }
            catch (System.ArgumentOutOfRangeException)
            {
                throw new ConfigDataException("L'indice de la valeur du type de mesure " + measureType.Name + " n'est pas correcte, veuillez la modifier.");
            }

            this.MeasureType = measureType;
        }

        public Measure(MeasureType measureType)
        {
            this.NominalValue = 0.0;
            this.TolerancePlus = 0.0;
            this.ToleranceMinus = 0.0;
            this.Value = 0.0;

            this.MeasureType = measureType;
        }
    }
}
