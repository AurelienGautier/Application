using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Win32;

namespace Application
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// La fenêtre est composée d'une barre de navigation et d'un control.
    /// Chaque control correspond à une fonctionnalité différente de l'application.
    /// </summary>
    public partial class MainWindow : Window
    {
        // Les différents controls de l'application. Ils correspondent chacun à une fonctionnalité.
        private readonly UI.UserControls.FillMitutoyoFormControl fillMitutoyoFormControl;
        private readonly UI.UserControls.FillAyonisFormControl fillAyonisFormControl;
        private readonly UI.UserControls.MeasureTypesControl measureTypesControl;
        private readonly UI.UserControls.AddMeasureType addMesureTypeControl;

        [DllImport("Kernel32")]
        public static extern void AllocConsole();

        public MainWindow()
        {
            AllocConsole();
            InitializeComponent();

            this.fillMitutoyoFormControl = new UI.UserControls.FillMitutoyoFormControl();
            this.fillAyonisFormControl = new UI.UserControls.FillAyonisFormControl();
            this.measureTypesControl = new UI.UserControls.MeasureTypesControl();
            this.addMesureTypeControl = new UI.UserControls.AddMeasureType();

            // Par défaut, on affiche le control de remplissage de formulaire Mitutoyo.
            CurrentControl.Content = this.fillMitutoyoFormControl;
        }

        /*-------------------------------------------------------------------------*/
        // Les méthodes suivantes servent à changer de control en fonction de l'action de l'utilisateur.

        private void goToFillMitutoyoForm(object sender, RoutedEventArgs e)
        {
            CurrentControl.Content = this.fillMitutoyoFormControl;
        }

        private void goToFillAyonisForm(object sender, RoutedEventArgs e)
        {
            CurrentControl.Content = this.fillAyonisFormControl;
        }

        public void goToMeasureTypes(object sender, RoutedEventArgs e)
        {
            this.measureTypesControl.BindData();
            CurrentControl.Content = this.measureTypesControl;
        }

        public void goToAddMeasureType(object sender, RoutedEventArgs e)
        {
            this.addMesureTypeControl.LoadMeasureType(null);
            CurrentControl.Content = this.addMesureTypeControl;
        }

        public void goToModifyMeasureType(Data.MeasureType measureType)
        {
            this.addMesureTypeControl.LoadMeasureType(measureType);
            CurrentControl.Content = this.addMesureTypeControl;
        }

        /*-------------------------------------------------------------------------*/

        /**
         * Permet de sélectionner un fichier d'image correspondant à la signature de l'utilisateur pour signer le formulaire.
         */
        private void chooseSignature(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();

            dialog.Filter = "(*.png)|*.png|(*.jpg)|*.jpg";
            dialog.Title = "Sélectionner une signature";

            String fileName = "";

            if (dialog.ShowDialog() == true) fileName = dialog.FileName;

            if(fileName == "") return;

            try
            {
                Data.ConfigSingleton.Instance.SetSignature(fileName);
            }
            catch (Exceptions.ConfigDataException ex)
            {
                DisplayError(ex.Message);
            }
        }

        /*-------------------------------------------------------------------------*/

        /**
         * DisplayError permet d'afficher une fenêtre d'erreur avec un message d'erreur donné.
         */
        public static void DisplayError(String errorMessage)
        {
            String caption = "Erreur";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Error;

            MessageBox.Show(errorMessage, caption, button, icon, MessageBoxResult.Yes);
        }

        /*-------------------------------------------------------------------------*/
    }
}