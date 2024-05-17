using Application.Data;
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

            Dictionary<string, string> pageNames = ConfigSingleton.Instance.GetPageNames();
            HeaderPage.Text = pageNames["HeaderPage"];
            MeasurePage.Text = pageNames["MeasurePage"];
        }

        private void saveSettingsClick(object sender, System.Windows.RoutedEventArgs e)
        {
            this.saveHeaderFields();

            this.savePageNames();

            this.displaySuccess("Les paramètres ont été sauvegardés avec succès");
        }

        private void saveHeaderFields()
        {
            if (Designation.Text == "" || PlanNb.Text == "" || Index.Text == "" || ClientName.Text == "")
            {
                this.displayError("Tous les champs d'en-tête doivent être remplis");
                return;
            }

            ConfigSingleton.Instance.SetHeaderFieldsMatch(Designation.Text, PlanNb.Text, Index.Text, ClientName.Text);
        }

        private void savePageNames()
        {
            if(HeaderPage.Text == "" || MeasurePage.Text == "")
            {
                this.displayError("Tous les noms de page doivent être remplis");
                return;
            }

            ConfigSingleton.Instance.SetPageNames(HeaderPage.Text, MeasurePage.Text);
        }

        private void displaySuccess(String sucessMessage)
        {
            MessageBox.Show(sucessMessage, "Succès", MessageBoxButton.OK, MessageBoxImage.None);
        }

        private void displayError(String errorMessage)
        {
            MessageBox.Show(errorMessage, "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
