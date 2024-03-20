using System.Text;
using System.Windows;
using NPOI.XWPF.UserModel;
using System.IO;
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

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            string templatePath = txtFilePath.Text;
            string fileName = Filler.GetTmpFileName("docx");
            string outputPath = Filler.GetOutPath(templatePath, fileName);

            //C:\Users\asus\Documents\GitHub\ZPP-Grupa7\sprawozdanie.docx

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
        }

        public void ShowDocumentPreview(string filePath)
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