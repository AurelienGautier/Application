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
    }
}
