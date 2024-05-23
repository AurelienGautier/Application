using Application.Data;
using Application.Exceptions;
using Application.Parser;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Represents a user control that allows the user to automatically fill a form with data got from a machine.
    /// </summary>
    public partial class FillFormControl : UserControl
    {
        private FormFillingManager formFillingManager;
        List<String> machines;
        ObservableCollection<Form> forms;

        private BindingList<ComboBoxItem> ComboBoxItems;
        private BindingList<String> AvailableOptions;

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Initializes a new instance of the FillFormControl class.
        /// </summary>
        public FillFormControl()
        {
            InitializeComponent();

            // Initialize the list of machines and bind it to the form
            this.machines = new List<String> { "Mitutoyo", "Ayonis" };
            Machines.ItemsSource = this.machines;
            Machines.SelectedIndex = 0;

            this.formFillingManager = new FormFillingManager();

            // Retrieve the list of existing forms
            this.forms = new ObservableCollection<Form>(ConfigSingleton.Instance.GetMitutoyoForms());
            Forms.ItemsSource = this.forms.Select(form => form.Name).ToList();
            Forms.SelectedIndex = 0;

            // Add the code attributes of each standards element to AvailableOptions
            List<String> standards = ConfigSingleton.Instance.GetStandards().Select(standard => standard.Code).ToList();
            AvailableOptions = new BindingList<string>(standards);

            ComboBoxItems = new BindingList<ComboBoxItem>();

            ComboBoxItem firstComboBox = new ComboBoxItem { AvailableOptions = AvailableOptions };
            firstComboBox.SelectedOption = AvailableOptions[0];

            ComboBoxItems.Add(firstComboBox);

            Standards.ItemsSource = ComboBoxItems;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Updates the data related to the standards.
        /// </summary>
        public void BindData()
        {
            List<String> standards = ConfigSingleton.Instance.GetStandards().Select(standard => standard.Code).ToList();
            AvailableOptions = new BindingList<string>(standards);

            foreach (var comboBoxItem in ComboBoxItems)
            {
                comboBoxItem.AvailableOptions = AvailableOptions;
            }

            Standards.ItemsSource = ComboBoxItems;
            Standards.Items.Refresh();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Action called when the user clicks decides to create a new form.
        /// </summary>
        private void fillAform(object sender, RoutedEventArgs e)
        {
            this.callFormFilling(null, SignForm.IsChecked == true, false);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Action called when the user decides to modify an existing form.
        /// </summary>
        private void modifyAform(object sender, RoutedEventArgs e)
        {
            String formToModify = this.formFillingManager.GetFileToOpen("Choose the form to modify", "(*.xlsx;*.xlsm)|*.xlsx;*.xlsm");
            if (formToModify == "") return;

            this.callFormFilling(formToModify, SignForm.IsChecked == true, true);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Prepares the Form object and sends it to the FormFillingManager to fill the form.
        /// </summary>
        private void callFormFilling(String? formToOverwritePath, bool sign, bool modify)
        {
            // Check the signature if the user wants to sign the document
            if (sign && ConfigSingleton.Instance.Signature == null)
            {
                MainWindow.DisplayError("It is impossible to sign this document because the signature is incorrect or not selected.");
                return;
            }

            // Find the selected form in the list of forms
            Form? form = this.forms.ToList<Form>().Find(f => f.Name == (String)Forms.SelectedItem);

            if (form == null)
            {
                MainWindow.DisplayError("The selected form is not supported.");
                return;
            }

            form.Modify = modify;
            form.Sign = sign;

            if (formToOverwritePath != null) form.Path = formToOverwritePath;

            List<Standard> standards = this.getStandardsFromComboBox();

            // Fill the form using the FormFillingManager
            this.formFillingManager.ManageFormFilling(form, this.getParser((String)Machines.SelectedItem), standards);
        }

        /*-------------------------------------------------------------------------*/

        private List<Standard> getStandardsFromComboBox()
        {
            List<Standard> standards = new List<Standard>();

            var selectedOptions = ComboBoxItems.Select(comboBoxItem => comboBoxItem.SelectedOption);

            foreach (var selectedOption in selectedOptions)
            {
                if (selectedOption == null)
                    throw new ConfigDataException("Oops, something went wrong!");

                Standard? standard = ConfigSingleton.Instance.GetStandardFromCode(selectedOption);
                if (standard == null) throw new ConfigDataException("The selected standard does not exist.");

                standards.Add(standard);
            }

            return standards;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Action called when the user selects a different machine.
        /// Changes the content of the forms list based on the selected machine.
        /// </summary>
        private void changeMachine(object sender, SelectionChangedEventArgs e)
        {
            if ((String)Machines.SelectedItem == "Ayonis")
                this.forms = new ObservableCollection<Form>(ConfigSingleton.Instance.GetAyonisForms());
            else this.forms = new ObservableCollection<Form>(ConfigSingleton.Instance.GetMitutoyoForms());

            Forms.ItemsSource = this.forms.Select(form => form.Name).ToList();

            if ((String)Forms.SelectedItem == null) Forms.SelectedIndex = 0;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Returns the parser corresponding to the selected machine.
        /// </summary>
        private Parser.Parser getParser(String selectedMachine)
        {
            if (selectedMachine == "Ayonis") return new ExcelParser();
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

    /// <summary>
    /// Represents an item in the dropdown list.
    /// </summary>
    public class ComboBoxItem
    {
        public BindingList<string>? AvailableOptions { get; set; }
        public string? SelectedOption { get; set; }
    }
}
