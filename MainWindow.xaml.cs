using Application.Exceptions;
using Application.Writers;
using Microsoft.Win32;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

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
    }
}