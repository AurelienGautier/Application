using Application.Data;
using Application.Exceptions;
using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Interaction logic for Settings.xaml
    /// </summary>
    public partial class Settings : UserControl
    {
        public Settings()
        {
            InitializeComponent();

            // Retrieve header fields from configuration
            Dictionary<string, string> headerFields = ConfigSingleton.Instance.GetHeaderFieldsMatch();

            // Fill header fields in the user interface
            if (headerFields.ContainsKey("Designation")) Designation.Text = headerFields["Designation"];
            if (headerFields.ContainsKey("PlanNb")) PlanNb.Text = headerFields["PlanNb"];
            if (headerFields.ContainsKey("Index")) Index.Text = headerFields["Index"];
            if (headerFields.ContainsKey("ClientName")) ClientName.Text = headerFields["ClientName"];
            if (headerFields.ContainsKey("ObservationNum")) ObservationNum.Text = headerFields["ObservationNum"];
            if (headerFields.ContainsKey("PieceReceptionDate")) PieceReceptionDate.Text = headerFields["PieceReceptionDate"];
            if (headerFields.ContainsKey("Observations")) Observations.Text = headerFields["Observations"];

            // Retrieve page names from configuration
            Dictionary<string, string> pageNames = ConfigSingleton.Instance.GetPageNames();
            if (pageNames.ContainsKey("HeaderPage")) HeaderPage.Text = pageNames["HeaderPage"];
            if (pageNames.ContainsKey("MeasurePage")) MeasurePage.Text = pageNames["MeasurePage"];
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Handle click event of the "Save Settings" button
        /// </summary>
        /// <param name="sender">The object that raised the event</param>
        /// <param name="e">The event arguments</param>
        private void saveSettingsClick(object sender, RoutedEventArgs e)
        {
            try
            {
                // Save header fields
                this.saveHeaderFields();
                // Save page names
                this.savePageNames();
                // Display success message
                this.displaySuccess("Settings have been saved successfully");
            }
            catch (InvalidFieldException ex)
            {
                // Display error message in case of exception
                this.displayError(ex.Message);
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Save header fields
        /// </summary>
        private void saveHeaderFields()
        {
            // Check if all header fields are filled
            if (Designation.Text == "" || PlanNb.Text == "" || Index.Text == "" || ClientName.Text == "" || ObservationNum.Text == "" || PieceReceptionDate.Text == "" || Observations.Text == "")
            {
                throw new InvalidFieldException("All header fields must be filled");
            }

            // Save header fields in the configuration
            ConfigSingleton.Instance.SetHeaderFieldsMatch(Designation.Text, PlanNb.Text, Index.Text, ClientName.Text, ObservationNum.Text, PieceReceptionDate.Text, Observations.Text);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Save page names
        /// </summary>
        private void savePageNames()
        {
            // Check if all page names are filled
            if (HeaderPage.Text == "" || MeasurePage.Text == "")
            {
                throw new InvalidFieldException("All page names must be filled");
            }

            // Save page names in the configuration
            ConfigSingleton.Instance.SetPageNames(HeaderPage.Text, MeasurePage.Text);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Update standards
        /// </summary>
        /// <param name="sender">The object that raised the event</param>
        /// <param name="e">The event arguments</param>
        private void updateStandards(object sender, RoutedEventArgs e)
        {
            try
            {
                // Update standards in the configuration
                ConfigSingleton.Instance.UpdateStandards();
                // Display success message
                this.displaySuccess("Standards have been updated successfully");
            }
            catch (ConfigDataException ex)
            {
                // Display error message in case of exception
                this.displayError(ex.Message);
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Display success message
        /// </summary>
        /// <param name="successMessage">The success message to display</param>
        private void displaySuccess(string successMessage)
        {
            MessageBox.Show(successMessage, "Success", MessageBoxButton.OK, MessageBoxImage.None);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Display error message
        /// </summary>
        /// <param name="errorMessage">The error message to display</param>
        private void displayError(string errorMessage)
        {
            MessageBox.Show(errorMessage, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        /*-------------------------------------------------------------------------*/
    }
}
