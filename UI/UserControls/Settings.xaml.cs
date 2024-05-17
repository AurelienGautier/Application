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
        String designation;
        String planNb;
        String index;
        String clientName;

        public Settings()
        {
            InitializeComponent();

            Dictionary<string, string> headerFields = ConfigSingleton.Instance.GetHeaderFieldsMatch();
            this.designation = headerFields["Designation"];
            this.planNb = headerFields["PlanNb"];
            this.index = headerFields["Index"];
            this.clientName = headerFields["ClientName"];

            Designation.Text = this.designation;
            PlanNb.Text = this.planNb;
            Index.Text = this.index;
            ClientName.Text = this.clientName;
        }

        private void saveSettingsClick(object sender, System.Windows.RoutedEventArgs e)
        {
            if (Designation.Text == "" || PlanNb.Text == "" || Index.Text == "" || ClientName.Text == "")
            {
                this.displayError("Tous les champs doivent être remplis");
                return;
            }

            ConfigSingleton.Instance.SetHeaderFieldsMatch(Designation.Text, PlanNb.Text, Index.Text, ClientName.Text);

            this.displaySuccess("Les champs d'en-tête ont été modifiées avec succès");
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
