using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Logique d'interaction pour MeasureTypesControl.xaml
    /// </summary>
    public partial class MeasureTypesControl : UserControl
    {
        public MeasureTypesControl()
        {
            InitializeComponent();

            var myList = Data.ConfigSingleton.Instance.GetMeasureTypes();

            MeasureTypes.ItemsSource = myList; 
        }

        private void modifyMeasureType(object sender, System.Windows.RoutedEventArgs e)
        {
            Button buttn = (Button)sender;
            String? libelle = buttn.Tag.ToString();
            if (libelle == null) return;

            Data.MeasureType? measureType = Data.ConfigSingleton.Instance.GetMeasureTypeFromLibelle(libelle);
            if(measureType == null) return;

            MainWindow parentWindow = (MainWindow)Window.GetWindow(this);
            parentWindow.goToModifyMeasureType(measureType);
        }

        private void deleteMeasureType(object sender, System.Windows.RoutedEventArgs e)
        {
            // To do
        }
    }
}
