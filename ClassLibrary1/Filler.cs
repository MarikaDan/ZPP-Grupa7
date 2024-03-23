using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
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
            "<grupa>",
            "<nr_cwiczenia>"
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

        public static IEnumerable<TagReplacement> PrepareReplacementsFromList(XWPFDocument document)
        {
            var result = new List<TagReplacement>();

            foreach (var para in document.Paragraphs)
            {
                PlaceholderTags.ForEach(tag =>
                {
                    if (para.ParagraphText.Contains(tag))
                    {
                        result.Add(new TagReplacement(tag));
                    }
                });
            }

            return result;
        }

        public static IEnumerable<TagReplacement> PrepareReplacementsFromRegex(XWPFDocument document)
        {
            var result = new List<TagReplacement>();

            var regex = new Regex(@"<[^<>]+>");
            foreach (var para in document.Paragraphs)
            {
                var matches = regex.Matches(para.ParagraphText);
                foreach (Match match in matches)
                {
                    result.Add(new TagReplacement(match.Value));
                }
            }

            return result;
        }

        public static void ReplacePlaceholders(XWPFDocument document, IEnumerable<TagReplacement> replacements)
        {
            foreach (var para in document.Paragraphs)
            {
                var tmpList = replacements.ToList();
                foreach (var replacement in tmpList)
                {
                    if (!para.ParagraphText.Contains(replacement.Tag)) continue;

                    if (replacement.Value is null) continue;

                    para.ReplaceText(replacement.Tag, replacement.Value);
                }
            }
        }
    }
}