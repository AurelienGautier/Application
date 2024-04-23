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
    public partial class MainWindow : System.Windows.Window
    {
        [DllImport("Kernel32")]
        public static extern void AllocConsole();

        [DllImport("Kernel32", SetLastError = true)]
        public static extern void FreeConsole();

        private FillFormControl fillFormControl;
        private MeasureTypesControl measureTypesControl;

        public MainWindow()
        {
            InitializeComponent();

            this.fillFormControl = new FillFormControl();
            this.measureTypesControl = new MeasureTypesControl();

            CurrentControl.Content = this.fillFormControl;
        }

        private void goToFillForm(object sender, RoutedEventArgs e)
        {
            CurrentControl.Content = this.fillFormControl;
        }

        private void goToMeasureTypes(object sender, RoutedEventArgs e)
        {
            CurrentControl.Content = this.measureTypesControl;
        }
    }
}