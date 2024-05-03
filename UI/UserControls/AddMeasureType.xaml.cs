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

        /*-------------------------------------------------------------------------*/

        public AddMeasureType()
        {
            InitializeComponent();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Permet de pré-remplir les champs du formulaire avec les informations du type de mesure à modifier
         * @param measureType Le type de mesure à charger
         */
        public void LoadMeasureType(Data.MeasureType? measureType)
        {
            this.measureType = measureType;

            if (measureType == null) return;

            TextBox measureName = (TextBox)this.FindName("MeasureName");
            measureName.Text = measureType.Name;

            TextBox measureNominalValueIndex = (TextBox)this.FindName("MeasureNominalValueIndex");
            measureNominalValueIndex.Text = measureType.NominalValueIndex.ToString();

            TextBox measureTolerancePlusIndex = (TextBox)this.FindName("MeasureTolerancePlusIndex");
            measureTolerancePlusIndex.Text = measureType.TolPlusIndex.ToString();

            TextBox measureValueIndex = (TextBox)this.FindName("MeasureValueIndex");
            measureValueIndex.Text = measureType.ValueIndex.ToString();

            TextBox measureToleranceMinusIndex = (TextBox)this.FindName("MeasureToleranceMinusIndex");
            measureToleranceMinusIndex.Text = measureType.TolMinusIndex.ToString();

            TextBox measureSymbol = (TextBox)this.FindName("MeasureSymbol");
            measureSymbol.Text = measureType.Symbol;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Permet de récupérer les informations du formulaire une fois rempli
         * @return Le type de mesure créé
         */
        public Data.MeasureType GetMeasureTypeFromPage()
        {
            TextBox measureName = (TextBox)this.FindName("MeasureName");
            TextBox measureNominalValueIndex = (TextBox)this.FindName("MeasureNominalValueIndex");
            TextBox measureTolerancePlusIndex = (TextBox)this.FindName("MeasureTolerancePlusIndex");
            TextBox measureValueIndex = (TextBox)this.FindName("MeasureValueIndex");
            TextBox measureToleranceMinusIndex = (TextBox)this.FindName("MeasureToleranceMinusIndex");
            TextBox measureSymbol = (TextBox)this.FindName("MeasureSymbol");

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
            Data.ConfigSingleton.Instance.UpdateMeasureType(this.measureType, this.GetMeasureTypeFromPage());

            MainWindow parentWindow = (MainWindow)Window.GetWindow(this);
            parentWindow.goToMeasureTypes(sender, e);
        }

        /*-------------------------------------------------------------------------*/
    }
}
