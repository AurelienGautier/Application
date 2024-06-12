using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Represents a user control for displaying and managing measure types.
    /// </summary>
    public partial class MeasureTypesControl : UserControl
    {
        /// <summary>
        /// Initializes a new instance of the MeasureTypesControl class.
        /// </summary>
        public MeasureTypesControl()
        {
            InitializeComponent();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Retrieves the list of measure types and displays them in the list.
        /// </summary>
        public void BindData()
        {
            // Retrieve the list of measure types
            List<Data.MeasureType> list = Data.ConfigSingleton.Instance.GetMeasureTypes();

            // Create an observable collection from the list
            ObservableCollection<Data.MeasureType> newItems = new ObservableCollection<Data.MeasureType>(list);

            // Set the items source of the MeasureTypes list to the new collection
            MeasureTypes.ItemsSource = newItems;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Handles the event when the user clicks the add measure type button.
        /// </summary>
        /// <param name="sender">The object that raised the event.</param>
        /// <param name="e">The event arguments.</param>
        private void addMeasureType(object sender, System.Windows.RoutedEventArgs e)
        {
            // Get the parent window
            MainWindow parentWindow = (MainWindow)Window.GetWindow(this);

            // Call the goToAddMeasureType method of the parent window
            parentWindow.goToAddMeasureType(sender, e);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Redirects to the measure type modification page.
        /// </summary>
        /// <param name="sender">The object that raised the event.</param>
        /// <param name="e">The event arguments.</param>
        private void modifyMeasureType(object sender, System.Windows.RoutedEventArgs e)
        {
            // Get the button that raised the event
            Button buttn = (Button)sender;

            // Get the libelle (label) from the button's tag
            String? libelle = buttn.Tag.ToString();

            if (libelle == null) return;

            // Get the measure type from the libelle
            Data.MeasureType? measureType = Data.ConfigSingleton.Instance.GetMeasureTypeFromLibelle(libelle);

            if (measureType == null) return;

            // Get the parent window
            MainWindow parentWindow = (MainWindow)Window.GetWindow(this);

            // Call the goToModifyMeasureType method of the parent window with the measure type as parameter
            parentWindow.goToModifyMeasureType(measureType);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Deletes a measure type selected by the user.
        /// </summary>
        /// <param name="sender">The object that raised the event.</param>
        /// <param name="e">The event arguments.</param>
        private void deleteMeasureType(object sender, System.Windows.RoutedEventArgs e)
        {
            // Get the button that raised the event
            Button button = (Button)sender;

            // Get the libelle (label) from the button's tag
            String? libelle = button.Tag.ToString();

            // If the libelle is null, return
            if (libelle == null) return;

            // Show a confirmation message box
            MessageBoxResult result = MessageBox.Show("Êtes-vous sûr de vouloir supprimer le type de mesure " + libelle + "?", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            // If the user clicks No, return
            if (result == MessageBoxResult.No) return;

            // Delete the measure type
            Data.ConfigSingleton.Instance.DeleteMeasureType(libelle);

            // Update the data binding
            this.BindData();
        }

        /*-------------------------------------------------------------------------*/
    }
}
