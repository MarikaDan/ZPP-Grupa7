using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using GrapeCity.Documents.Word;
using System.Drawing;
using GrapeCity.Documents.Word.Layout;


namespace ZPP_1
{
    internal class Program
    {
        static List<string> placeHolderDictionary = new List<string>();

        static void Main(string[] args)
        {
            placeHolderDictionary.Add("<imie>");
            placeHolderDictionary.Add("<data>");

            var tmp = @"D:\PLIKI\STUDIA\SEM6\ZPP\Projekt1\Template1.docx";


            using (var rs = File.OpenRead(tmp))
            {
                var generateFile = @"D:\PLIKI\STUDIA\SEM6\ZPP\Projekt1\output1.docx";

                using (var doc = new XWPFDocument(rs))
                {
                    foreach (var para in doc.Paragraphs)
                    {
                        foreach (var placeholder in placeHolderDictionary)
                        {
                            if (para.ParagraphText.Contains(placeholder))
                            {
                                para.ReplaceText(placeholder, "przykładowy tekst");
                            }
                        }
                    }
                    using (var ws = File.Create(generateFile))
                    {
                        doc.Write(ws);


                    }

                }

            }

            var wordDoc = new GcWordDocument();
            var path = Path.Combine(@"D:\PLIKI\STUDIA\SEM6\ZPP\Projekt1\output1.docx");
            wordDoc.Load(path);
            using (var layout = new GcWordLayout(wordDoc))
            {
                PdfOutputSettings pdfOutput = new PdfOutputSettings();
                pdfOutput.CompressionLevel = CompressionLevel.Fastest;
                pdfOutput.ConformanceLevel = GrapeCity.Documents.Pdf.PdfAConformanceLevel.PdfA1a;
                pdfOutput.BackColor = Color.Transparent;
                pdfOutput.DocumentInfo = new GrapeCity.Documents.Pdf.DocumentInfo()
                {
                    Author = "Test",
                    Title = "TestPDF",
                    CreationDate = new GrapeCity.Documents.Pdf.PdfDateTime(DateTime.Now)
                };
                layout.SaveAsPdf(@"D:\PLIKI\STUDIA\SEM6\ZPP\Projekt1\output1.pdf", null, pdfOutput);
            }

        }
    }
}