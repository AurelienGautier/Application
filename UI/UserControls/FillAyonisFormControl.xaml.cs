using Application.Parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

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
            switch (Forms.SelectedItem)
            {
                case "Rapport 1 pièce":
                    this.formFillingManager.FullOnePieceFile(30, Environment.CurrentDirectory + "\\form\\rapport1piece", 26, new ExcelParser());
                    break;
                case "Rapport 5 pièces":
                    this.formFillingManager.FullFivePieesFile(new ExcelParser());
                    break;
            }
        }
    }
}
