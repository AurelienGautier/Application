using Application.Exceptions;
using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Logique d'interaction pour AddMesureType.xaml
    /// Ce control a pour objectif de permettre à l'utilisateur de créer un type de mesure ou d'en modifier un déjà existant
    /// </summary>
    public partial class AddMeasureType : UserControl
    {
        // Le type de mesure à modifier (dans le cas d'une modification). Il prend la valeur null dans le cas d'une création
        private Data.MeasureType? measureType;

        // Les différents champs à remplir pour créer un type de mesure
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

            this.measureName = (TextBox)this.FindName("MeasureName");
            this.measureNominalValueIndex = (TextBox)this.FindName("MeasureNominalValueIndex");
            this.measureTolerancePlusIndex = (TextBox)this.FindName("MeasureTolerancePlusIndex");
            this.measureValueIndex = (TextBox)this.FindName("MeasureValueIndex");
            this.measureToleranceMinusIndex = (TextBox)this.FindName("MeasureToleranceMinusIndex");
            this.measureSymbol = (TextBox)this.FindName("MeasureSymbol");
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Permet de pré-remplir les champs du formulaire avec les informations du type de mesure à modifier
         * @param measureType Le type de mesure à charger
         */
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

        /**
         * Permet de récupérer les informations du formulaire une fois rempli
         * @return Le type de mesure créé
         */
        public Data.MeasureType GetMeasureTypeFromPage()
        {
            if (this.measureName.Text == "" 
                || this.measureNominalValueIndex.Text == "" 
                || this.measureTolerancePlusIndex.Text == "" 
                || this.measureValueIndex.Text == "" 
                || this.measureToleranceMinusIndex.Text == "" 
                || this.measureSymbol.Text == ""
            )
                throw new ConfigDataException("Tous les champs doivent être remplis");

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

        /**
         * Permet de sauvegarder le type de mesure créé ou modifié
         */
        private void saveMeasureType(object sender, RoutedEventArgs e)
        {
            try
            {
                Data.ConfigSingleton.Instance.UpdateMeasureType(this.measureType, this.GetMeasureTypeFromPage());

                MainWindow parentWindow = (MainWindow)Window.GetWindow(this);
                parentWindow.goToMeasureTypes(sender, e);
            }
            catch(ConfigDataException ex)
            {
                MainWindow.DisplayError(ex.Message);
            }
        }

        /*-------------------------------------------------------------------------*/
    }
}
