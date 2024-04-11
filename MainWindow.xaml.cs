﻿using Application.Exceptions;
using Application.Writers;
using Microsoft.Win32;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

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

        public MainWindow()
        {
            InitializeComponent();
        }

        public void Rapport1Piece_Click(object sender, RoutedEventArgs e)
        {
            String fileToParse = this.getFileToOpen();
            if (fileToParse == "") return;
            String fileToSave = this.getFileToSave();
            if (fileToSave == "") return;

            try
            {
                Parser parser = new Parser();
                List<Piece> data = parser.ParseFile(fileToParse);

                OnePieceWriter excelWriter = new OnePieceWriter(fileToSave);
                excelWriter.WriteData(data);
            }
            catch(IncorrectFormatException)
            {
                this.displayError("Le format du fichier est incorrect.");
            }
            catch(ExcelFileAlreadyInUse)
            {
                this.displayError("Le fichier excel est déjà en cours d'utilisation");
            }
        }

        private void Rapport5Pieces_Click(object sender, RoutedEventArgs e)
        {
            /*AllocConsole();*/

            var dialog = new OpenFolderDialog();
            String folderName = "";

            // Récupération du dossier contenant les fichiers à traiter
            if(dialog.ShowDialog() == true)
            {
                folderName = dialog.FolderName;
            }

            if (folderName == "") return;

            List<String> files = new List<String>();

            DirectoryInfo directory = new DirectoryInfo(folderName);

            // Récupération de la liste des fichiers du répertoire
            if(directory.Exists)
            {
                foreach(FileInfo file in directory.GetFiles())
                {
                    files.Add(file.FullName);
                }
            }

            Parser parser = new Parser();

            List<Piece> data = new List<Piece>();

            // Traitement de chaque fichier
            foreach(String file in files)
            {
                try
                {
                    List<Piece> pieces = parser.ParseFile(file);
                    data.AddRange(pieces);
                }
                catch(IncorrectFormatException)
                {
                    this.displayError("Le format du fichier " + file + " est incorrect.");
                }
            }

            FivePiecesWriter excelWriter = new FivePiecesWriter(this.getFileToSave());
            excelWriter.WriteData(data);

            /*FreeConsole();*/
        }

        private String getFileToOpen()
        {
            var dialog = new OpenFileDialog();
            dialog.FileName = "Document";
            dialog.DefaultExt = ".txt";

            String fileName = "";

            if (dialog.ShowDialog() == true)
            {
                fileName = dialog.FileName;
            }

            return fileName;
        }

        private String getFileToSave()
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Fichiers Excel (*.xlsx)|*.xlsx";
            saveFileDialog.FileName = "rappport1piece";

            String fileName = "";
         
            if (saveFileDialog.ShowDialog() == true)
            {
                fileName = saveFileDialog.FileName;
            }

            if(fileName.Length > 5)
                fileName = fileName.Remove(fileName.Length - 5);

            return fileName;
        }

        private void displayError(String errorMessage)
        {
            String caption = "Erreur";
            MessageBoxButton button = MessageBoxButton.OK;
            MessageBoxImage icon = MessageBoxImage.Error;
            MessageBoxResult result;

            result = MessageBox.Show(errorMessage, caption, button, icon, MessageBoxResult.Yes);
        }
    }
}