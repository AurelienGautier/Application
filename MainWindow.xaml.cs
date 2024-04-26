using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Win32;

namespace Application
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("Kernel32")]
        public static extern void AllocConsole();

        [DllImport("Kernel32", SetLastError = true)]
        public static extern void FreeConsole();

        private UI.UserControls.FillMitutoyoFormControl fillMitutoyoFormControl;
        private UI.UserControls.FillAyonisFormControl fillAyonisFormControl;
        private UI.UserControls.MeasureTypesControl measureTypesControl;
        private UI.UserControls.AddMesureType addMesureTypeControl;

        public MainWindow()
        {
            AllocConsole();

            InitializeComponent();

            this.fillMitutoyoFormControl = new UI.UserControls.FillMitutoyoFormControl();
            this.fillAyonisFormControl = new UI.UserControls.FillAyonisFormControl();
            this.measureTypesControl = new UI.UserControls.MeasureTypesControl();
            this.addMesureTypeControl = new UI.UserControls.AddMesureType();

            CurrentControl.Content = this.fillMitutoyoFormControl;
        }

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

        private void chooseSignature(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();

            String fileName = "";

            if (dialog.ShowDialog() == true) fileName = dialog.FileName;

            Data.ConfigSingleton.Instance.SetSignature(fileName);
        }

        public void goToModifyMeasureType(Data.MeasureType measureType)
        {
            this.addMesureTypeControl.LoadMeasureType(measureType);
            CurrentControl.Content = this.addMesureTypeControl;
        }
    }
}