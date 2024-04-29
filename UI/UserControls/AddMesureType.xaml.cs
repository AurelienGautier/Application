using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Logique d'interaction pour AddMesureType.xaml
    /// </summary>
    public partial class AddMesureType : UserControl
    {
        private Data.MeasureType? measureType;

        public AddMesureType()
        {
            InitializeComponent();
        }

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

        public Data.MeasureType GetMeasureType()
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

        private void saveMeasureType(object sender, RoutedEventArgs e)
        {
            Data.ConfigSingleton.Instance.UpdateMeasureType(this.measureType, this.GetMeasureType());

            MainWindow parentWindow = (MainWindow)Window.GetWindow(this);
            parentWindow.goToMeasureTypes(sender, e);
        }
    }
}
