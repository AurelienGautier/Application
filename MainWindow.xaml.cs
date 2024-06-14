using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;

namespace Application
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// The window is composed of 2 parts: a navbar and a main control. 
    /// Each control contains a page corresponding to a different functionality of the application.
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly UI.UserControls.FillFormControl fillFormControl;
        private readonly UI.UserControls.MeasureTypesControl measureTypesControl;
        private readonly UI.UserControls.AddMeasureTypeControl addMesureTypeControl;
        private readonly UI.UserControls.Settings settingsControl;

        private ImageSource? logo = null;

        private bool measureTypesWarning = false;
        private bool settingsWarning = false;


        /// <summary>
        /// Initializes a new instance of the MainWindow class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            this.fillFormControl = new UI.UserControls.FillFormControl();
            this.measureTypesControl = new UI.UserControls.MeasureTypesControl();
            this.addMesureTypeControl = new UI.UserControls.AddMeasureTypeControl();
            this.settingsControl = new UI.UserControls.Settings();

            try
            {
                logo = new BitmapImage(new System.Uri(Environment.CurrentDirectory + "\\res\\lelogodefoula.png"));
                Logo.Source = logo;
            }
            catch
            {
                // Do nothing, the logo will just not be displayed
            }

            CurrentControl.Content = this.fillFormControl;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Change the control to display the fill form control.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        private void goToFillForm(object sender, RoutedEventArgs e)
        {
            this.fillFormControl.BindData();
            CurrentControl.Content = this.fillFormControl;
        }

        /// <summary>
        /// Change the control to display the measure types control.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        public void goToMeasureTypes(object sender, RoutedEventArgs e)
        {
            if (!this.measureTypesWarning)
            {
                DisplayWarning("Attention, la modification des types de mesures peut entraîner des erreurs dans des fichiers corrects. Ne modifiez les types de mesures que si vous savez ce que vous faites.");
                this.measureTypesWarning = true;
            }

            this.measureTypesControl.BindData();
            CurrentControl.Content = this.measureTypesControl;
        }

        /// <summary>
        /// Change the control to display the add measure type control.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        public void goToAddMeasureType(object sender, RoutedEventArgs e)
        {
            this.addMesureTypeControl.LoadMeasureType(null);
            CurrentControl.Content = this.addMesureTypeControl;
        }

        /// <summary>
        /// Change the control to display the modify measure type control.
        /// The control is the same as the add measure type control, but the fields are pre-filled with the data of the measure type to modify (passed as a parameter).
        /// </summary>
        /// <param name="measureType">The measure type to modify.</param>
        public void goToModifyMeasureType(Data.MeasureType measureType)
        {
            this.addMesureTypeControl.LoadMeasureType(measureType);
            CurrentControl.Content = this.addMesureTypeControl;
        }

        /// <summary>
        /// Change the control to display the settings control.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        public void goToSettings(object sender, RoutedEventArgs e)
        {
            if (!this.settingsWarning)
            {
                DisplayWarning("Attention, la modification des paramètres peut entraîner des erreurs dans les formulaires remplis. Ne modifiez les paramètres que si vous savez ce que vous faites.");
                this.settingsWarning = true;
            }

            CurrentControl.Content = this.settingsControl;
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Allows the user to select an image file corresponding to the user's signature to sign the form.
        /// </summary>
        /// <param name="sender">The object that triggered the event.</param>
        /// <param name="e">The event arguments.</param>
        private void chooseSignature(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();

            dialog.Filter = "(*.png;*.jpg)|*.png;*.jpg";
            dialog.Title = "Sélectionner une signature";

            String filePath = "";

            if (dialog.ShowDialog() == true) filePath = dialog.FileName;

            if (filePath == "") return;

            try
            {
                Data.ConfigSingleton.Instance.SetSignature(filePath);
            }
            catch (Exceptions.ConfigDataException ex)
            {
                DisplayError(ex.Message);
            }
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Displays an error window with the given error message.
        /// </summary>
        /// <param name="errorMessage">The error message to display.</param>
        public static void DisplayError(String errorMessage)
        {
            String caption = "Erreur";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Error;

            MessageBox.Show(errorMessage, caption, button, icon, MessageBoxResult.Yes);
        }

        /*-------------------------------------------------------------------------*/

        /// <summary>
        /// Displays a warning window with the given warning message.
        /// </summary>
        /// <param name="warningMessage">The warning message to display.</param>
        public static void DisplayWarning(String warningMessage)
        {
            String caption = "Avertissement";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Warning;

            MessageBox.Show(warningMessage, caption, button, icon, MessageBoxResult.Yes);
        }

        /*-------------------------------------------------------------------------*/
    }
}