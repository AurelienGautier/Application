using Application.Exceptions;
using Application.Writers;
using Microsoft.Win32;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;

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
            this.FullOnePieceFile(30, Environment.CurrentDirectory + "\\form\\rapport1piece", 26, 66);
        }

        public void OutillageControle_Click(object sender, RoutedEventArgs e)
        {
            this.FullOnePieceFile(26, Environment.CurrentDirectory + "\\form\\outillageDeControle", 25, 62);
        }

        public void FullOnePieceFile(int firstLine, String formPath, int designLine, int operatorLine)
        {
            String fileToParse = this.getFileToOpen();
            if (fileToParse == "") return;
            String fileToSave = this.getFileToSave();
            if (fileToSave == "") return;

            try
            {
                Parser parser = new Parser();
                List<Piece> data = parser.ParseFile(fileToParse);
                Dictionary<string, string> header = parser.GetHeader();

                OnePieceWriter excelWriter = new OnePieceWriter(fileToSave, firstLine, formPath);
                excelWriter.WriteHeader(header, designLine);
                excelWriter.WriteData(data);
            }
            catch(MeasureTypeNotFoundException)
            {
                this.displayError("Un type de mesure n'a pas été reconnu dans le fichier " + fileToParse);
            }
            catch (IncorrectFormatException)
            {
                this.displayError("Le format du fichier est incorrect.");
            }
            catch (ExcelFileAlreadyInUseException)
            {
                this.displayError("Le fichier excel est déjà en cours d'utilisation");
            }
        }

        public void Rapport5Pieces_Click(object sender, RoutedEventArgs e)
        {
            String folderName = this.getFolderToOpen();
            if (folderName == "") return;

            DirectoryInfo directory = new DirectoryInfo(folderName);
            if (!directory.Exists) return;

            Parser parser = new Parser();
            List<Piece> data = new List<Piece>();

            // Parsing de tous les fichiers du répertoire
            foreach (FileInfo file in directory.GetFiles())
            {
                try
                {
                    data.AddRange(parser.ParseFile(file.FullName));
                }
                catch (IncorrectFormatException)
                {
                    this.displayError("Le format du fichier " + file.FullName + " est incorrect.");
                    return;
                }
                catch(MeasureTypeNotFoundException)
                {
                    this.displayError("Un type de mesure n'a pas été trouvé dans le fichier " + file.FullName);
                    return;
                }
            }
            Dictionary<string, string> header = parser.GetHeader();

            String fileToSave = this.getFileToSave();
            if (fileToSave == "") return;

            try
            {
                FivePiecesWriter excelWriter = new FivePiecesWriter(fileToSave);
                excelWriter.WriteHeader(header, 25);
                excelWriter.WriteData(data);
            }
            catch (ExcelFileAlreadyInUseException)
            {
                this.displayError("Le fichier excel est déjà en cours d'utilisation");
            }
        }

        private String getFileToOpen()
        {
            var dialog = new OpenFileDialog();

            String fileName = "";

            if (dialog.ShowDialog() == true) fileName = dialog.FileName;

            return fileName;
        }

        private String getFolderToOpen()
        {
            var dialog = new OpenFolderDialog();

            String folderName = "";

            if (dialog.ShowDialog() == true) folderName = dialog.FolderName;

            return folderName;
        }

        private String getFileToSave()
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Fichiers Excel (*.xlsx)|*.xlsx";
            saveFileDialog.FileName = "rapport";

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

            MessageBox.Show(errorMessage, caption, button, icon, MessageBoxResult.Yes);
        }
    }
}