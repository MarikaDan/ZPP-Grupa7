using System.Windows;
using ClassLibrary1;

namespace ZPP_1_UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string _templatePath;

        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                DefaultExt = ".docx",
                Filter = "Word documents (.docx)|*.docx"
            };

            var result = dialog.ShowDialog();

            if (result != true) return;
            _templatePath = dialog.FileName;
            selectedPath.Text = _templatePath;
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(_templatePath))
            {
                MessageBox.Show("Wybierz plik szablonu.");
                return;
            }

            var fileName = Filler.GetTmpFileName("docx");
            var docxTmpPath = Filler.GetOutPath(_templatePath, fileName);

            var document = Filler.OpenDocument(_templatePath);

            var replacements = Filler.PrepareReplacements(document);

            //Do testów, trzeba zrobić to w pętli albo zrobić to jakoś ładniej
            replacements["<imie>"] = txtImie.Text;
            replacements["<data>"] = txtData.Text;

            Filler.ReplacePlaceholders(document, replacements);

            Filler.SaveDocxFile(document, docxTmpPath);

            var pdfFile = Filler.GetOutPath(_templatePath, Filler.GetTmpFileName("pdf"));
            Filler.ConvertToPdf(docxTmpPath, pdfFile);

            Filler.RemoveTmpFile(docxTmpPath);

            MessageBox.Show("Plik został wygenerowany.");

            Filler.OpenWithDefaultProgram(pdfFile);
        }

    }
}