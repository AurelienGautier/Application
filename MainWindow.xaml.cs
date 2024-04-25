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

        public MainWindow()
        {
            AllocConsole();

            InitializeComponent();

            fillMitutoyoFormControl = new UI.UserControls.FillMitutoyoFormControl();
            fillAyonisFormControl = new UI.UserControls.FillAyonisFormControl();
            measureTypesControl = new UI.UserControls.MeasureTypesControl();

            CurrentControl.Content = fillMitutoyoFormControl;
        }

        private void goToFillMitutoyoForm(object sender, RoutedEventArgs e)
        {
            CurrentControl.Content = fillMitutoyoFormControl;
        }

        private void goToFillAyonisForm(object sender, RoutedEventArgs e)
        {
            CurrentControl.Content = fillAyonisFormControl;
        }

        private void goToMeasureTypes(object sender, RoutedEventArgs e)
        {
            CurrentControl.Content = measureTypesControl;
        }

        private void chooseSignature(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();

            String fileName = "";

            if (dialog.ShowDialog() == true) fileName = dialog.FileName;

            Data.ConfigSingleton.Instance.Signature = fileName;
        }
    }
}