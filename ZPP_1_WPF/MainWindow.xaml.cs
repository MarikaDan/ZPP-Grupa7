using System.Collections.ObjectModel;
using System.Windows;
using ClassLibrary1;
using NPOI.XWPF.UserModel;

namespace ZPP_1_UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _templatePath;
        private XWPFDocument _document;

        private ObservableCollection<TagReplacement> _tagReplacements;
        public ObservableCollection<TagReplacement> TagReplacements => _tagReplacements;

        public MainWindow()
        {
            InitializeComponent();

            _tagReplacements = new ObservableCollection<TagReplacement>();
            DataContext = this;

            tagList.ItemsSource = Filler.PlaceholderTags;
        }

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

        private void PrepareFile()
        {
            if (string.IsNullOrEmpty(_templatePath))
            {
                MessageBox.Show("Wybierz plik szablonu.");
                return;
            }

            _document = Filler.OpenDocument(_templatePath);

            var tags = UseRegexCheckBox.IsChecked switch
            {
                true => Filler.PrepareReplacementsFromRegex(_document),
                _ => Filler.PrepareReplacementsFromList(_document)
            };

            _tagReplacements = new ObservableCollection<TagReplacement>(tags);
            replacementsGrid.ItemsSource = _tagReplacements;
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            var fileName = Filler.GetTmpFileName("docx");
            var docxTmpPath = Filler.GetOutPath(_templatePath, fileName);


            Filler.ReplacePlaceholders(_document, TagReplacements);
            Filler.SaveDocxFile(_document, docxTmpPath);

            var pdfFile = Filler.GetOutPath(_templatePath, Filler.GetTmpFileName("pdf"));
            Filler.ConvertToPdf(docxTmpPath, pdfFile);

            Filler.RemoveTmpFile(docxTmpPath);

            MessageBox.Show("Plik został wygenerowany.");

            Filler.OpenWithDefaultProgram(pdfFile);
        }

        private void OpenFile_Click(object sender, RoutedEventArgs e) => PrepareFile();
    }

}