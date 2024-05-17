using Application.Data;
using Application.Exceptions;
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

        public ObservableCollection<ComboBoxItem> ComboBoxItems { get; }
        public ObservableCollection<String> AvailableOptions { get; }

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

            // Ajout des attributs code de chaque élément de standards à AvailableOptions
            AvailableOptions = new ObservableCollection<string>(ConfigSingleton.Instance.GetStandards().Select(standard => standard.Code));

            ComboBoxItems = new ObservableCollection<ComboBoxItem>();

            ComboBoxItem firstComboBox = new ComboBoxItem { AvailableOptions = AvailableOptions };
            firstComboBox.SelectedOption = AvailableOptions[0];

            ComboBoxItems.Add(firstComboBox);

            Standards.ItemsSource = ComboBoxItems;
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
            // Ajouter la récupération des étalons

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

            List<Standard> standards = this.getStandardsFromComboBox();

            // Remplissage du formulaire en utilisant le FormFillingManager
            this.formFillingManager.ManageFormFilling(form, this.getParser((String) Machines.SelectedItem), standards);
        }

        /*-------------------------------------------------------------------------*/

        private List<Standard> getStandardsFromComboBox()
        {
            List<Standard> standards = new List<Standard>();

            var selectedOptions = ComboBoxItems.Select(comboBoxItem => comboBoxItem.SelectedOption);

            foreach (var selectedOption in selectedOptions)
            {
                if (selectedOption == null)
                    throw new ConfigDataException("Nan mé watzeufeuk ct pa sencé spacé ssa !!!");

                Standard? standard = ConfigSingleton.Instance.GetStandardFromCode(selectedOption);
                if (standard == null) throw new ConfigDataException("L'étalon sélectionné n'existe pas.");

                standards.Add(standard);
            }

            return standards;
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

            if((String)Forms.SelectedItem == null) Forms.SelectedIndex = 0;
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

        /*-------------------------------------------------------------------------*/

        private void AddStandard_Click(object sender, RoutedEventArgs e)
        {
            ComboBoxItem comboBoxItem = new ComboBoxItem { AvailableOptions = AvailableOptions };
            comboBoxItem.SelectedOption = AvailableOptions[0];
            ComboBoxItems.Add(comboBoxItem);
        }

        /*-------------------------------------------------------------------------*/

        private void RemoveStandard_Click(object sender, RoutedEventArgs e)
        {
            Button button = (Button)sender;
            ComboBoxItem optionToRemove = (ComboBoxItem)button.DataContext;
            ComboBoxItems.Remove(optionToRemove);
        }

        /*-------------------------------------------------------------------------*/
    }

    public class ComboBoxItem
    {
        public ObservableCollection<string>? AvailableOptions { get; set; }
        public string? SelectedOption { get; set; }
    }
}
