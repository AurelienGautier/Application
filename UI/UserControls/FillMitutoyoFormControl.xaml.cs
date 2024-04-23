using Application.Exceptions;
using Application.Parser;
using Application.Writers;
using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace Application.UI.UserControls
{
    /// <summary>
    /// Logique d'interaction pour FillForm.xaml
    /// </summary>
    public partial class FillMitutoyoFormControl : UserControl
    {
        FormFillingManager formFillingManager;

        public FillMitutoyoFormControl()
        {
            InitializeComponent();

            this.formFillingManager = new FormFillingManager();

            List<string> items = new List<string> {
                "Rapport 1 pièce",
                "Outillage de contrôle",
                "Rapport 5 pièces",
                "Bague lisse",
                "Calibre à machoire",
                "Capabilité",
                "Étalon de colonne de mesure",
                "Tampon lisse"
            };

            Forms.ItemsSource = items;
        }

        private void FillAform(object sender, RoutedEventArgs e)
        {
            switch (Forms.SelectedItem)
            {
                case "Rapport 1 pièce":
                    this.formFillingManager.FullOnePieceFile(30, Environment.CurrentDirectory + "\\form\\rapport1piece", 26, new TextFileParser());
                    break;
                case "Outillage de contrôle":
                    this.formFillingManager.FullOnePieceFile(26, Environment.CurrentDirectory + "\\form\\outillageDeControle", 25, new TextFileParser());
                    break;
                case "Rapport 5 pièces":
                    this.formFillingManager.FullFivePieesFile(new TextFileParser());
                    break;
            }
        }
    }
}
