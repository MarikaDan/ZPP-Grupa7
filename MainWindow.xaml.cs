using System.Windows;
using NPOI.XWPF.UserModel;
using System.IO;
using System.Text;
using System.Windows.Xps.Packaging;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            string templatePath = @"ścieżka_do_szablonu.docx"; // Zmień na właściwą ścieżkę
            string outputPath = @"ścieżka_do_wyjściowego_pliku.docx"; // Zmień na właściwą ścieżkę

            using (var rs = File.OpenRead(templatePath))
            {
                using (var doc = new XWPFDocument(rs))
                {
                    foreach (var para in doc.Paragraphs)
                    {
                        if (para.ParagraphText.Contains("<imie>"))
                        {
                            para.ReplaceText("<imie>", txtImie.Text);
                        }
                        if (para.ParagraphText.Contains("<data>"))
                        {
                            para.ReplaceText("<data>", txtData.Text);
                    }

                    using (var ws = File.Create(outputPath))
                    {
                        doc.Write(ws);
                    }
                }
            }

            MessageBox.Show("Plik został wygenerowany.");

            // Wyświetl podgląd dokumentu
            ShowDocumentPreview(outputPath);
        }

        private void ShowDocumentPreview(string filePath)
        {
            StringBuilder htmlContent = new StringBuilder();
            htmlContent.AppendLine("<html>");
            htmlContent.AppendLine("<head></head>");
            htmlContent.AppendLine("<body>");

            using (var stream = File.OpenRead(filePath))
            {
                var doc = new XWPFDocument(stream);
                foreach (var paragraph in doc.Paragraphs)
                {
                    htmlContent.AppendLine("<p>" + paragraph.ParagraphText + "</p>");
                }
            }

            htmlContent.AppendLine("</body>");
            htmlContent.AppendLine("</html>");

            // Ustaw zawartość HTML dla WebBrowser
            webBrowser.NavigateToString(htmlContent.ToString());
        }
    }
}
