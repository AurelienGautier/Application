using Application.Data;
using Application.Exceptions;
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
        readonly private FormFillingManager formFillingManager;
        readonly private List<MeasureMachine> machines;
        private ObservableCollection<Form> forms;

        readonly private BindingList<ComboBoxItem> ComboBoxItems;
        private BindingList<String> AvailableOptions;

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Initializes a new instance of the FillFormControl class.
        /// </summary>
        public FillFormControl()
        {
            InitializeComponent();

            // Initialize the list of machines and bind it to the form
            this.machines = ConfigSingleton.Instance.Machines;
            Machines.ItemsSource = this.machines.Select(machine => machine.Name).ToList();
            Machines.SelectedIndex = 0;

            // Initialize the FormFillingManager
            this.formFillingManager = new FormFillingManager();

            // Retrieve the list of existing forms
            this.forms = new ObservableCollection<Form>(this.machines[0].PossiblesForms);
            Forms.ItemsSource = this.forms.Select(form => form.Name).ToList();
            Forms.SelectedIndex = 0;

            // Add the code attributes of each standards element to AvailableOptions
            List<String> standards = ConfigSingleton.Instance.GetStandards().Select(standard => standard.Code).ToList();
            AvailableOptions = new BindingList<string>(standards);

            // Initialize the ComboBoxItems list
            ComboBoxItems = [];
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
            try
            {
                // Get data entered by the user
                MeasureMachine machine = this.getCorrectMachine();
                Form form = this.getCorrectForm(machine);
                form.SourceFiles = this.getCorrectSourceFiles();
                form.DestinationPath = this.getCorrectDestinationPath();
                form.Standards = this.getStandardsFromComboBox();
                form.Modify = Modify.IsChecked == true;

                // Check if the signature is correct
                if (SignForm.IsChecked == true)
                {
                    checkSignature();
                    form.Sign = true;
                }

                // If the user wants to modify a form, ask him to select the form to modify
                if (form.Modify)
                {
                    String formPathToModify;
                    formPathToModify = FormFillingManager.GetFileToOpen("Choisir le formulaire à modifier", "(*.xlsx;*.xlsm)|*.xlsx;*.xlsm");
                    if (formPathToModify == "") return;

                    form.Path = formPathToModify;
                }

                this.formFillingManager.ManageFormFilling(form, DestinationPathTextBox.Text);
            }
            catch (InvalidFieldException ex)
            {
                MainWindow.DisplayError(ex.Message);
            }
            catch (ConfigDataException ex)
            {
                MainWindow.DisplayError(ex.Message);
            }
        }

        /*-------------------------------------------------------------------------*/

        private MeasureMachine getCorrectMachine()
        {
            MeasureMachine? machine = this.machines.Find(m => m.Name == (String)Machines.SelectedItem);

            return machine ?? throw new InvalidFieldException("La machine sélectionnée n'existe pas");
        }

        /*-------------------------------------------------------------------------*/

        private Form getCorrectForm(MeasureMachine machine)
        {
            Form? form = machine.PossiblesForms.Find(f => f.Name == (String)Forms.SelectedItem) ?? throw new InvalidFieldException("Le formulaire sélectionné n'existe pas pour la machine " + machine.Name);

            if (form.Type == FormType.Capability)
            {
                if (MeasureNum.Text == "")
                {
                    throw new InvalidFieldException("Veuillez renseigner le/les numéro(s) de mesure pour la capabilité");
                }

                try
                {
                    List<String> list = [.. MeasureNum.Text.Split(',')];
                    List<int> capabilityValues = list.Select(int.Parse).ToList();

                    form.CapabilityMeasureNumber = capabilityValues;
                }
                catch
                {
                    throw new InvalidFieldException("Les numéros de mesure doivent être des nombres");
                }
            }

            return form;
        }

        /*-------------------------------------------------------------------------*/

        private List<String> getCorrectSourceFiles()
        {
            if (SourcePathTextBox.Text == "")
            {
                throw new InvalidFieldException("Veuillez renseigner le chemin du/des fichier(s) ou du dossier source");
            }

            List<String> sourceFiles = [.. SourcePathTextBox.Text.Split('|')];

            foreach (String sourceFile in sourceFiles)
            {
                if (!File.Exists(sourceFile))
                {
                    throw new InvalidFieldException("Le fichier source " + sourceFile + " n'existe pas");
                }
            }

            return sourceFiles;
        }

        /*-------------------------------------------------------------------------*/

        private string getCorrectDestinationPath()
        {
            if (DestinationPathTextBox.Text == "")
            {
                throw new InvalidFieldException("Veuillez renseigner le chemin du formulaire de destination");
            }

            if (!Directory.Exists(Path.GetDirectoryName(DestinationPathTextBox.Text)))
            {
                throw new InvalidFieldException("Le chemin du rapport de destination est inaccessible");
            }

            return DestinationPathTextBox.Text;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Get the list of standards selected by the user.
        /// </summary>
        /// <returns></returns>
        /// <exception cref="ConfigDataException"></exception>
        private List<Standard> getStandardsFromComboBox()
        {
            List<Standard> standards = [];

            var selectedOptions = ComboBoxItems.Select(comboBoxItem => comboBoxItem.SelectedOption);

            foreach (var selectedOption in selectedOptions)
            {
                if (selectedOption == null)
                    throw new ConfigDataException("Waw mé cé pa neaurmal ssa ia 1 preaublaym atancion oulala");

                Standard? standard = ConfigSingleton.Instance.GetStandardFromCode(selectedOption) ?? throw new ConfigDataException("L'étalon sélectionné n'existe pas.");

                standards.Add(standard);
            }

            return standards;
        }

        /*-------------------------------------------------------------------------*/

        private static void checkSignature()
        {
            if (ConfigSingleton.Instance.Signature == null)
            {
                throw new ConfigDataException("Il est impossible de signer le document car la signature est incorrecte ou non sélectionée");
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Action called when the user selects a different machine.
        /// Changes the content of the forms list based on the selected machine.
        /// </summary>
        private void changeMachine(object sender, SelectionChangedEventArgs e)
        {
            var selectedMachine = this.machines.Find(machine => machine.Name == (String)Machines.SelectedItem);

            if (selectedMachine == null) return;

            this.forms = new ObservableCollection<Form>(selectedMachine.PossiblesForms);

            Forms.ItemsSource = this.forms.Select(form => form.Name).ToList();

            if ((String)Forms.SelectedItem == null) Forms.SelectedIndex = 0;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Action called when the user selects a different form.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void changeForm(object sender, SelectionChangedEventArgs e)
        {
            Form? form = this.forms.ToList<Form>().Find(f => f.Name == (String)Forms.SelectedItem);

            if (form == null)
            {
                MainWindow.DisplayError("Le formulaire sélectionné n'existe pas");
                return;
            }

            if (form.Type == FormType.Capability)
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
        /// Action called when the user wants to add a new standard to the list.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AddStandard_Click(object sender, RoutedEventArgs e)
        {
            ComboBoxItem comboBoxItem = new()
            {
                AvailableOptions = AvailableOptions,
                SelectedOption = AvailableOptions[0]
            };
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
            Form? selectedForm = this.forms.ToList<Form>().Find(f => f.Name == (String)Forms.SelectedItem);
            if (selectedForm == null) return;

            List<String> filesToParse = this.formFillingManager.GetFilesToOpen("Choisir le fichier à convertir", selectedForm.MeasureMachine.Parser.GetFileExtension(), selectedForm.DataFrom == DataFrom.Files);
            
            if (filesToParse.Count == 0) return;

            SourcePathTextBox.Text = "";

            foreach (String file in filesToParse)
            {
                SourcePathTextBox.Text += file + "|";
            }

            SourcePathTextBox.Text = SourcePathTextBox.Text.Remove(SourcePathTextBox.Text.Length - 1);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Open a dialog to select the folder to parse
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void browseSourceFolder(object sender, RoutedEventArgs e)
        {
            String folderToParse = FormFillingManager.GetFolderToOpen("Choisir le dossier à convertir");
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
            String fileToSave = FormFillingManager.GetFileToSave();
            if (fileToSave == "") return;

            DestinationPathTextBox.Text = fileToSave;
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
