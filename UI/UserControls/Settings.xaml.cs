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

            Designation.Text = headerFields["Designation"];
            PlanNb.Text = headerFields["PlanNb"];
            Index.Text = headerFields["Index"];
            ClientName.Text = headerFields["ClientName"];
            ObservationNum.Text = headerFields["ObservationNum"];
            PieceReceptionDate.Text = headerFields["PieceReceptionDate"];
            Observations.Text = headerFields["Observations"];

            Dictionary<string, string> pageNames = ConfigSingleton.Instance.GetPageNames();
            HeaderPage.Text = pageNames["HeaderPage"];
            MeasurePage.Text = pageNames["MeasurePage"];
        }

        /*-------------------------------------------------------------------------*/

        private void saveSettingsClick(object sender, System.Windows.RoutedEventArgs e)
        {
            this.saveHeaderFields();

            this.savePageNames();

            this.displaySuccess("Les paramètres ont été sauvegardés avec succès");
        }

        /*-------------------------------------------------------------------------*/

        private void saveHeaderFields()
        {
            if (Designation.Text == "" || PlanNb.Text == "" || Index.Text == "" || ClientName.Text == "" || ObservationNum.Text == "" || PieceReceptionDate.Text == "" || Observations.Text == "")
            {
                this.displayError("Tous les champs d'en-tête doivent être remplis");
                return;
            }

            ConfigSingleton.Instance.SetHeaderFieldsMatch(Designation.Text, PlanNb.Text, Index.Text, ClientName.Text, ObservationNum.Text, PieceReceptionDate.Text, Observations.Text);
        }

        /*-------------------------------------------------------------------------*/

        private void savePageNames()
        {
            if(HeaderPage.Text == "" || MeasurePage.Text == "")
            {
                this.displayError("Tous les noms de page doivent être remplis");
                return;
            }

            ConfigSingleton.Instance.SetPageNames(HeaderPage.Text, MeasurePage.Text);
        }

        /*-------------------------------------------------------------------------*/

        private void updateStandards(object sender, System.Windows.RoutedEventArgs e)
        {
            try
            {
                ConfigSingleton.Instance.UpdateStandards();
            }
            catch(ConfigDataException ex)
            {
                this.displayError(ex.Message);
                return;
            }

            this.displaySuccess("Les standards ont été mis à jour avec succès");
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
