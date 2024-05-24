using Application.Data;
using Application.Exceptions;
using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Logique d'interaction pour Settings.xaml
    /// </summary>
    public partial class Settings : UserControl
    {
        public Settings()
        {
            InitializeComponent();

            Dictionary<string, string> headerFields = ConfigSingleton.Instance.GetHeaderFieldsMatch();

            if (headerFields.ContainsKey("Designation")) Designation.Text = headerFields["Designation"];
            if (headerFields.ContainsKey("PlanNb")) PlanNb.Text = headerFields["PlanNb"];
            if (headerFields.ContainsKey("Index")) Index.Text = headerFields["Index"];
            if (headerFields.ContainsKey("ClientName")) ClientName.Text = headerFields["ClientName"];
            if (headerFields.ContainsKey("ObservationNum")) ObservationNum.Text = headerFields["ObservationNum"];
            if (headerFields.ContainsKey("PieceReceptionDate")) PieceReceptionDate.Text = headerFields["PieceReceptionDate"];
            if (headerFields.ContainsKey("Observations")) Observations.Text = headerFields["Observations"];

            Dictionary<string, string> pageNames = ConfigSingleton.Instance.GetPageNames();
            if(pageNames.ContainsKey("HeaderPage")) HeaderPage.Text = pageNames["HeaderPage"];
            if(pageNames.ContainsKey("MeasurePage")) MeasurePage.Text = pageNames["MeasurePage"];
        }

        /*-------------------------------------------------------------------------*/

        private void saveSettingsClick(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                this.saveHeaderFields();
                this.savePageNames();
                this.displaySuccess("Les paramètres ont été sauvegardés avec succès");
            }
            catch (InvalidFieldException ex)
            {
                this.displayError(ex.Message);
            }
        }

        /*-------------------------------------------------------------------------*/

        private void saveHeaderFields()
        {
            if (Designation.Text == "" || PlanNb.Text == "" || Index.Text == "" || ClientName.Text == "" || ObservationNum.Text == "" || PieceReceptionDate.Text == "" || Observations.Text == "")
            {
                throw new InvalidFieldException("Tous les champs d'en-tête doivent être remplis");
            }

            ConfigSingleton.Instance.SetHeaderFieldsMatch(Designation.Text, PlanNb.Text, Index.Text, ClientName.Text, ObservationNum.Text, PieceReceptionDate.Text, Observations.Text);
        }

        /*-------------------------------------------------------------------------*/

        private void savePageNames()
        {
            if(HeaderPage.Text == "" || MeasurePage.Text == "")
            {
                throw new InvalidFieldException("Tous les noms de page doivent être remplis");
            }

            ConfigSingleton.Instance.SetPageNames(HeaderPage.Text, MeasurePage.Text);
        }

        /*-------------------------------------------------------------------------*/

        private void updateStandards(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                ConfigSingleton.Instance.UpdateStandards();
                this.displaySuccess("Les standards ont été mis à jour avec succès");
            }
            catch(ConfigDataException ex)
            {
                this.displayError(ex.Message);
            }
        }

        /*-------------------------------------------------------------------------*/

        private void displaySuccess(String sucessMessage)
        {
            MessageBox.Show(sucessMessage, "Succès", MessageBoxButton.OK, MessageBoxImage.None);
        }

        /*-------------------------------------------------------------------------*/

        private void displayError(String errorMessage)
        {
            MessageBox.Show(errorMessage, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        /*-------------------------------------------------------------------------*/
    }
}
