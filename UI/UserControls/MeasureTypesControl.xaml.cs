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
            // To do
        }

        private void deleteMeasureType(object sender, System.Windows.RoutedEventArgs e)
        {
            // To do
        }
    }
}
