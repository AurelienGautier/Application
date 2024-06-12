using Application.Data;
using Application.Exceptions;
using Application.Parser;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Interaction logic for FillFormControl.xaml
    /// </summary>
    public partial class FillFormControl : UserControl
    {
        private FormFillingManager formFillingManager;
        List<String> machines;
        ObservableCollection<Form> forms;

        private BindingList<ComboBoxItem> ComboBoxItems;
        private BindingList<String> AvailableOptions;

        private Form? currentForm = null;

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

            // Initialize the FormFillingManager
            this.formFillingManager = new FormFillingManager();

            // Retrieve the list of existing forms
            this.forms = new ObservableCollection<Form>(ConfigSingleton.Instance.GetMitutoyoForms());
            Forms.ItemsSource = this.forms.Select(form => form.Name).ToList();
            Forms.SelectedIndex = 0;

            this.currentForm = this.forms[0];

            // Add the code attributes of each standards element to AvailableOptions
            List<String> standards = ConfigSingleton.Instance.GetStandards().Select(standard => standard.Code).ToList();
            AvailableOptions = new BindingList<string>(standards);

            // Initialize the ComboBoxItems list
            ComboBoxItems = new BindingList<ComboBoxItem>();
            Standards.ItemsSource = ComboBoxItems;

            // Hide the measure number stack by default
            MeasureNumStack.Visibility = Visibility.Collapsed;
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
            if (!isFormCorrectlyFilled()) return;

            String? formToModify = null;

            if (this.currentForm == null) return;

            if (this.currentForm.Modify)
            {
                formToModify = this.formFillingManager.GetFileToOpen("Choisir le formulaire à modifier", "(*.xlsx;*.xlsm)|*.xlsx;*.xlsm");
                if (formToModify == "") return;
            }

            this.callFormFilling(formToModify);
            this.currentForm = this.forms.ToList<Form>().Find(f => f.Name == (String)Forms.SelectedItem);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Checks if the form is correctly filled.
        /// </summary>
        /// <returns></returns>
        private bool isFormCorrectlyFilled()
        {
            if (this.currentForm == null)
            {
                MainWindow.DisplayError("Le formulaire sélectionné n'est pas pris en compte.");
                return false;
            }

            // Check the signature if the user wants to sign the document
            if (SignForm.IsChecked == true && ConfigSingleton.Instance.Signature == null)
            {
                MainWindow.DisplayError("Il est impossible de signer le document car la signature est incorrecte ou non sélectionée.");
                return false;
            }

            if (SourcePathTextBox.Text == "")
            {
                MainWindow.DisplayError("Veuillez renseigner le chemins du fichier ou du dossier source.");
                return false;
            }

            if (DestinationPathTextBox.Text == "")
            {
                MainWindow.DisplayError("Veuillez renseigner le chemin du formulaire de destination.");
                return false;
            }
            else if (!Directory.Exists(Path.GetDirectoryName(DestinationPathTextBox.Text)))
            {
                MainWindow.DisplayError("Le chemin du dossier de destination n'existe pas.");
                return false;
            }

            if (this.currentForm.Type == FormType.Capability)
            {
                if (MeasureNum.Text == "")
                {
                    MainWindow.DisplayError("Veuillez renseigner le numéro de mesure.");
                    return false;
                }

                try
                {
                    List<String> list = MeasureNum.Text.Split(',').ToList();
                    List<int> capabilityValues = list.Select(int.Parse).ToList();

                    this.currentForm.CapabilityMeasureNumber = capabilityValues;
                }
                catch
                {
                    MainWindow.DisplayError("Le numéro de mesure doit être un nombre.");
                    return false;
                }
            }


            return true;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Prepares the Form object and sends it to the FormFillingManager to fill the form.
        /// </summary>
        private void callFormFilling(String? formToOverwritePath)
        {
            if (this.currentForm == null)
            {
                MainWindow.DisplayError("Le formulaire sélectionné n'est pas pris en compte.");
                return;
            }

            this.currentForm.Sign = SignForm.IsChecked == true;

            if (formToOverwritePath != null) this.currentForm.Path = formToOverwritePath;

            List<Standard> standards = this.getStandardsFromComboBox();

            // Fill the form using the FormFillingManager
            this.formFillingManager.ManageFormFilling(this.currentForm, this.getParser(), standards, DestinationPathTextBox.Text);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Get the list of standards selected by the user.
        /// </summary>
        /// <returns></returns>
        /// <exception cref="ConfigDataException"></exception>
        private List<Standard> getStandardsFromComboBox()
        {
            List<Standard> standards = new List<Standard>();

            var selectedOptions = ComboBoxItems.Select(comboBoxItem => comboBoxItem.SelectedOption);

            foreach (var selectedOption in selectedOptions)
            {
                if (selectedOption == null)
                    throw new ConfigDataException("Waw mé cé pa neaurmal ssa ia 1 preaublaym atancion oulala");

                Standard? standard = ConfigSingleton.Instance.GetStandardFromCode(selectedOption);
                if (standard == null) throw new ConfigDataException("L'étalon sélectionné n'existe pas.");

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

            this.changeForm(sender, e);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Action called when the user selects a different form.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void changeForm(object sender, SelectionChangedEventArgs e)
        {
            this.currentForm = this.forms.ToList<Form>().Find(f => f.Name == (String)Forms.SelectedItem);

            if (this.currentForm == null)
            {
                this.currentForm = this.forms[0];
            }

            if (this.currentForm.Type == FormType.Capability)
            {
                MeasureNumStack.Visibility = Visibility.Visible;
            }
            else
            {
                MeasureNumStack.Visibility = Visibility.Collapsed;
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Returns the parser corresponding to the selected machine.
        /// </summary>
        private Parser.Parser getParser()
        {
            if ((String)Machines.SelectedItem == "Ayonis") return new ExcelParser();
            return new TextFileParser();
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Action called when the user wants to add a new standard to the list.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddStandard_Click(object sender, RoutedEventArgs e)
        {
            ComboBoxItem comboBoxItem = new ComboBoxItem { AvailableOptions = AvailableOptions };
            comboBoxItem.SelectedOption = AvailableOptions[0];
            ComboBoxItems.Add(comboBoxItem);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Action called when the user wants to remove a standard from the list.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RemoveStandard_Click(object sender, RoutedEventArgs e)
        {
            Button button = (Button)sender;
            ComboBoxItem optionToRemove = (ComboBoxItem)button.DataContext;
            ComboBoxItems.Remove(optionToRemove);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Open a dialog box to select the file to parse.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void browseSourceFiles(object sender, RoutedEventArgs e)
        {
            if(this.currentForm == null) return;

            List<String> filesToParse = this.formFillingManager.GetFilesToOpen("Choisir le fichier à convertir", this.getParser().GetFileExtension(), this.currentForm.DataFrom == DataFrom.Files);
            
            if (filesToParse.Count == 0) return;

            foreach (String file in filesToParse)
            {
                SourcePathTextBox.Text += file + ";";
            }

            this.currentForm.SourceFiles = filesToParse;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Open a dialog to select the folder to parse
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void browseSourceFolder(object sender, RoutedEventArgs e)
        {
            String folderToParse = this.formFillingManager.GetFolderToOpen("Choisir le dossier à convertir");
            if (folderToParse == "") return;

            SourcePathTextBox.Text = folderToParse;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Action called when the user wants to select the path where to save the filled excel form.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void browseDestinationFile(object sender, RoutedEventArgs e)
        {
            String fileToSave = this.formFillingManager.GetFileToSave();
            if (fileToSave == "") return;

            DestinationPathTextBox.Text = fileToSave;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Action called when the user wants to modify an existing form.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Modify_Checked(object sender, RoutedEventArgs e)
        {
            if (this.currentForm == null) return;

            this.currentForm.Modify = true;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Action called when the user wants to create a new form.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void New_Check(object sender, RoutedEventArgs e)
        {
            if (this.currentForm == null) return;

            this.currentForm.Modify = false;
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
