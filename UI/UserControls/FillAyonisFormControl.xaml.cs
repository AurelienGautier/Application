using Application.Parser;
using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Logique d'interaction pour FillAyonisForm.xaml
    /// </summary>
    public partial class FillAyonisFormControl : UserControl
    {
        private FormFillingManager formFillingManager;

        public FillAyonisFormControl()
        {
            InitializeComponent();

            this.formFillingManager = new FormFillingManager();

            List<string> items = new List<string> {
                "Rapport 1 pièce",
                "Rapport 5 pièces",
            };

            Forms.ItemsSource = items;
        }

        private void FillAform(object sender, RoutedEventArgs e)
        {
            bool signForm = true;

            switch (Forms.SelectedItem)
            {
                case "Rapport 1 pièce":
                    this.formFillingManager.FullOnePieceFile(30, Environment.CurrentDirectory + "\\form\\rapport1piece", 26, new ExcelParser(), signForm);
                    break;
                case "Rapport 5 pièces":
                    this.formFillingManager.FullFivePieesFile(new ExcelParser(), signForm);
                    break;
            }
        }
    }
}
