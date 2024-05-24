using Application.Exceptions;
using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Interaction logic for AddMeasureType.xaml
    /// This control aims to allow the user to create a measurement type or modify an existing one
    /// </summary>
    public partial class AddMeasureType : UserControl
    {
        // The measurement type to modify (in case of modification). It takes the value null in case of creation
        private Data.MeasureType? measureType;

        // The different fields to fill in to create a measurement type
        private TextBox measureName;
        private TextBox measureNominalValueIndex;
        private TextBox measureTolerancePlusIndex;
        private TextBox measureValueIndex;
        private TextBox measureToleranceMinusIndex;
        private TextBox measureSymbol;

        /*-------------------------------------------------------------------------*/

        public AddMeasureType()
        {
            InitializeComponent();

            // The different fields to fill in to create a measurement type
            this.measureName = (TextBox)this.FindName("MeasureName");
            this.measureNominalValueIndex = (TextBox)this.FindName("MeasureNominalValueIndex");
            this.measureTolerancePlusIndex = (TextBox)this.FindName("MeasureTolerancePlusIndex");
            this.measureValueIndex = (TextBox)this.FindName("MeasureValueIndex");
            this.measureToleranceMinusIndex = (TextBox)this.FindName("MeasureToleranceMinusIndex");
            this.measureSymbol = (TextBox)this.FindName("MeasureSymbol");
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Allows pre-filling the form fields with information from the measurement type to modify
        /// </summary>
        /// <param name="measureType">The measurement type to load</param>
        public void LoadMeasureType(Data.MeasureType? measureType)
        {
            this.measureType = measureType;

            if (measureType == null)
            {
                measureName.Text = "";
                measureNominalValueIndex.Text = "";
                measureTolerancePlusIndex.Text = "";
                measureValueIndex.Text = "";
                measureToleranceMinusIndex.Text = "";
                measureSymbol.Text = "";
                return;
            }

            measureName.Text = measureType.Name;
            measureNominalValueIndex.Text = measureType.NominalValueIndex.ToString();
            measureTolerancePlusIndex.Text = measureType.TolPlusIndex.ToString();
            measureValueIndex.Text = measureType.ValueIndex.ToString();
            measureToleranceMinusIndex.Text = measureType.TolMinusIndex.ToString();
            measureSymbol.Text = measureType.Symbol;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Retrieves the form information once filled
        /// </summary>
        /// <returns>The created measurement type</returns>
        public Data.MeasureType GetMeasureTypeFromPage()
        {
            if (this.measureName.Text == ""
                || this.measureNominalValueIndex.Text == ""
                || this.measureTolerancePlusIndex.Text == ""
                || this.measureValueIndex.Text == ""
                || this.measureToleranceMinusIndex.Text == ""
            )
                throw new ConfigDataException("Tous les champs obligatoires doivent être remplis.");

            int tryInt;

            if (!int.TryParse(measureNominalValueIndex.Text, out tryInt)
                || !int.TryParse(measureTolerancePlusIndex.Text, out tryInt)
                || !int.TryParse(measureValueIndex.Text, out tryInt)
                || !int.TryParse(measureToleranceMinusIndex.Text, out tryInt)
            )
                throw new ConfigDataException("Tous les indices doivent être des nombres entiers.");

            return new Data.MeasureType()
            {
                Name = measureName.Text,
                NominalValueIndex = int.Parse(measureNominalValueIndex.Text),
                TolPlusIndex = int.Parse(measureTolerancePlusIndex.Text),
                ValueIndex = int.Parse(measureValueIndex.Text),
                TolMinusIndex = int.Parse(measureToleranceMinusIndex.Text),
                Symbol = measureSymbol.Text
            };
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Saves the created or modified measurement type
        /// </summary>
        /// <param name="sender">The object that triggered the event</param>
        /// <param name="e">The event arguments</param>
        private void saveMeasureType(object sender, RoutedEventArgs e)
        {
            try
            {
                Data.ConfigSingleton.Instance.UpdateMeasureType(this.measureType, this.GetMeasureTypeFromPage());

                MainWindow parentWindow = (MainWindow)Window.GetWindow(this);
                parentWindow.goToMeasureTypes(sender, e);
            }
            catch (ConfigDataException ex)
            {
                MainWindow.DisplayError(ex.Message);
            }
        }

        /*-------------------------------------------------------------------------*/
    }
}
