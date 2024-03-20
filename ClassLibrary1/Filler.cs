using System.Diagnostics;
using System.Linq;
using GrapeCity.Documents.Word;
using GrapeCity.Documents.Word.Layout;
using NPOI.XWPF.UserModel;

namespace ClassLibrary1
{
    public class Filler
    {
        public static List<string> PlaceholderTags = new List<string>
        {
            "<imie>",
            "<data>",
            "<przedmiot>",
            "<prowadzacy>",
            "<temat>",
            "<kierunek>",
            "<grupa>"
        };

        public static XWPFDocument OpenDocument(string path)
        {
            using var rs = File.OpenRead(path);
            var doc = new XWPFDocument(rs);

            return doc;
        }

        public static void SaveDocxFile(XWPFDocument document, string path)
        {
            using var ws = File.Create(path);
            document.Write(ws);
        }

        public static void RemoveTmpFile(string path)
        {
            if (!File.Exists(path)) return;
            File.Delete(path);
        }

        public static void ConvertToPdf(string docxPath, string outputPath)
        {
            var wordDoc = new GcWordDocument();
            wordDoc.Load(docxPath);

            using var layout = new GcWordLayout(wordDoc);
            var pdfOutput = new PdfOutputSettings();
            layout.SaveAsPdf(outputPath, null, pdfOutput);
        }
        public static string GetOutPath(string inputPath, string fileName)
        {
            var dir = Path.GetDirectoryName(inputPath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var outPath = Path.Combine(dir, fileName);
            return outPath;
        }

        public static string GetTmpFileName(string extension)
        {
            var tmp = Path.GetRandomFileName();
            tmp = Path.GetFileNameWithoutExtension(tmp);

            return $"{tmp}.{extension}";
        }

        public static void OpenWithDefaultProgram(string path)
        {
            using var fileOpener = new Process();

            fileOpener.StartInfo.FileName = "explorer";
            fileOpener.StartInfo.Arguments = "\"" + path + "\"";
            fileOpener.Start();
        }

        public static Dictionary<string, string> PrepareReplacements(XWPFDocument document)
        {
            var result = new Dictionary<string, string>();

            foreach (var para in document.Paragraphs)
            {
                PlaceholderTags.ForEach(tag =>
                {
                    if (para.ParagraphText.Contains(tag))
                    {
                        result.Add(tag, "");
                    }
                });
            }

            return result;
        }

        public static void ReplacePlaceholders(XWPFDocument document, Dictionary<string, string> replacements)
        {
            foreach (var para in document.Paragraphs)
            {
                foreach (var keyValuePair in replacements)
                {
                    if (!para.ParagraphText.Contains(keyValuePair.Key)) continue;

                    para.ReplaceText(keyValuePair.Key, keyValuePair.Value);
                }
            }
        }
    }
}