using Application.Data;
using Application.Parser;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Logique d'interaction pour FillForm.xaml
    /// Permet à l'utilisateur de remplir automatiquement un formulaire dont les données proviennent de la machine 
    /// </summary>
    public partial class FillFormControl : UserControl
    {
        private FormFillingManager formFillingManager;
        List<String> machines;
        ObservableCollection<Form> forms;

        /*-------------------------------------------------------------------------*/

        public FillFormControl()
        {
            InitializeComponent();

            // Initialisation de la liste des machines et binding avec le formulaire
            this.machines = new List<String> { "Mitutoyo", "Ayonis" };
            Machines.ItemsSource = this.machines;
            Machines.SelectedIndex = 0;

            this.formFillingManager = new FormFillingManager();

            // Récupération de la liste des formulaires existants
            this.forms = new ObservableCollection<Form>(ConfigSingleton.Instance.GetMitutoyoForms());
            Forms.ItemsSource = this.forms.Select(form => form.Name).ToList();
            Forms.SelectedIndex = 0;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * L'action qui est appelée lorsque l'utilisateur clique sur "nouveau"
         */
        private void fillAform(object sender, RoutedEventArgs e)
        {
            this.callFormFilling(null, SignForm.IsChecked == true, false);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * L'action qui est appelée lorsque l'utilisateur clique sur "modifier"
         */
        private void modifyAform(object sender, RoutedEventArgs e)
        {
            String formToModify = this.formFillingManager.GetFileToOpen("Choisir le formulaire à modifier", "(*.xlsx;*.xlsm)|*.xlsx;*.xlsm");
            if (formToModify == "") return;

            this.callFormFilling(formToModify, SignForm.IsChecked == true, true);
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Prépare l'objet de type Form et l'envoie au FormFillingManager pour remplir le formulaire
         */
        private void callFormFilling(String? formToOverwritePath, bool sign, bool modify)
        {
            // Vérification de la signature si l'utilisateur souhaite signer le document
            if (sign && ConfigSingleton.Instance.Signature == null)
            {
                MainWindow.DisplayError("Il est impossible de signer ce document car la signature est incorrect ou non sélectionnée.");
                return;
            }

            // Recherche du formulaire sélectionné dans la liste des formulaires
            Form? form = this.forms.ToList<Form>().Find(f => f.Name == (String)Forms.SelectedItem);

            if (form == null)
            {
                MainWindow.DisplayError("Le formulaire sélectionné n'est pas pris en charge.");
                return;
            }

            form.Modify = modify;
            form.Sign = sign;

            if (formToOverwritePath != null) form.Path = formToOverwritePath;

            // Remplissage du formulaire en utilisant le FormFillingManager
            this.formFillingManager.ManageFormFilling(form, this.getParser((String) Machines.SelectedItem));
        }

        /*-------------------------------------------------------------------------*/

        /**
         * L'action qui est appelée lorsque l'utilisateur sélectionne une autre machine
         * Change le contenu de la liste des formulaires en fonction de la machine sélectionnée
         */
        private void changeMachine(object sender, SelectionChangedEventArgs e)
        {
            if ((String)Machines.SelectedItem == "Ayonis") 
                this.forms = new ObservableCollection<Form>(ConfigSingleton.Instance.GetAyonisForms());
            else this.forms = new ObservableCollection<Form>(ConfigSingleton.Instance.GetMitutoyoForms());

            Forms.ItemsSource = this.forms.Select(form => form.Name).ToList();
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Retourne le parser correspondant à la machine sélectionnée
         */
        private Parser.Parser getParser(String selectedMachine)
        {
            if(selectedMachine == "Ayonis") return new ExcelParser();
            return new TextFileParser();
        }

        private void Forms_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        /*-------------------------------------------------------------------------*/
    }
}
