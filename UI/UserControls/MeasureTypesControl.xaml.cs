using System.Collections.ObjectModel;
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
        }

        /*-------------------------------------------------------------------------*/
        
        /**
         * Récupère la liste des types de mesures et les affiche dans la liste
         */
        public void BindData()
        {
            List<Data.MeasureType> list = Data.ConfigSingleton.Instance.GetMeasureTypes();
            ObservableCollection<Data.MeasureType> newItems = new ObservableCollection<Data.MeasureType>(list);

            MeasureTypes.ItemsSource = newItems;
        }

        /*-------------------------------------------------------------------------*/

        private void addMeasureType(object sender, System.Windows.RoutedEventArgs e)
        {
            MainWindow parentWindow = (MainWindow)Window.GetWindow(this);
            parentWindow.goToAddMeasureType(sender, e);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Redirige vers la page de modification du type de mesure
         */
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

        /*-------------------------------------------------------------------------*/

        /**
         * Supprime un type de mesure sélectionné par un utilisateur
         */
        private void deleteMeasureType(object sender, System.Windows.RoutedEventArgs e)
        {
            Button button = (Button)sender;
            String? libelle = button.Tag.ToString();
            if (libelle == null) return;

            MessageBoxResult result = MessageBox.Show("Êtes-vous sûr de vouloir supprimer le type de mesure " + libelle + " ?", "Avertissement", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.No) return;

            Data.ConfigSingleton.Instance.DeleteMeasureType(libelle);
            this.BindData();
        }

        /*-------------------------------------------------------------------------*/
    }
}
