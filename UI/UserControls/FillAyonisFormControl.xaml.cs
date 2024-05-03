using Application.Data;
using Application.Parser;
using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Logique d'interaction pour FillAyonisForm.xaml
    /// Permet à l'utilisateur de remplir automatiquement un formulaire dont les données proviennent de la machine Ayonis
    /// </summary>
    public partial class FillAyonisFormControl : UserControl
    {
        private FormFillingManager formFillingManager;
        List<Data.Form> forms;

        /*-------------------------------------------------------------------------*/

        public FillAyonisFormControl()
        {
            InitializeComponent();

            this.formFillingManager = new FormFillingManager();

            this.forms = ConfigSingleton.Instance.GetAyonisForms();

            Forms.ItemsSource = this.forms.Select(form => form.Name).ToList();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Permet de remplir un formulaire vierge automatiquement
         */
        private void fillAform(object sender, RoutedEventArgs e)
        {
            this.callFormFilling(null, new ExcelParser(), SignForm.IsChecked == true, false);
        }

        /*-------------------------------------------------------------------------*/

        private void modifyAform(object sender, RoutedEventArgs e)
        {
            String formToModify = this.formFillingManager.GetFileToOpen("Choisir le formulaire à modifier", "(*.xlsx;*.xlsm)|*.xlsx;*.xlsm");
            if (formToModify == "") return;

            this.callFormFilling(formToModify, new ExcelParser(), SignForm.IsChecked == true, true);
        }

        /*-------------------------------------------------------------------------*/

        private void callFormFilling(String? formToOverwritePath, Parser.Parser parser, bool sign, bool modify)
        {
            // Vérification de la signature
            if (sign && ConfigSingleton.Instance.Signature == null)
            {
                MainWindow.DisplayError("Il est impossible de signer ce document car la signature est incorrect ou non sélectionnée.");
                return;
            }

            // Recherche du formulaire sélectionné dans la liste des formulaires
            Form? form = this.forms.Find(f => f.Name == (String)Forms.SelectedItem);

            if (form == null)
            {
                MainWindow.DisplayError("Impossible de trouver le formulaire sélectionné.");
                return;
            }

            form.Modify = modify;
            form.Sign = sign;

            if (formToOverwritePath != null) form.Path = formToOverwritePath;

            // Remplissage du formulaire en utilisant le FormFillingManager
            switch (Forms.SelectedItem)
            {
                case "Rapport 1 pièce":
                    this.formFillingManager.FillOnePieceFile(form, parser);
                    break;
                case "Rapport 5 pièces":
                    this.formFillingManager.FillFivePiecesFile(form.Path, parser, sign, modify);
                    break;
            }
        }

        /*-------------------------------------------------------------------------*/
    }
}
