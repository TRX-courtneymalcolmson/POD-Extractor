using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace POD_Extractor
{
    class PdfOperations
    {

        public static List<String> SplitPdfPages(String sourcePdfPath, String outputDirectory)
        {
            // Convert word doc to pdf if required

            // INIT //
            PdfCopy copy = default;
            PdfReader reader = new PdfReader(sourcePdfPath);
            String savepath = default;

            List<String> splitPages = new List<String>();

            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                //create Document object
                Document document = new Document();
                savepath = outputDirectory +@"\"+ i + ".pdf";
                copy = new PdfCopy(document, new FileStream(savepath, FileMode.Create));
                splitPages.Add(savepath);


                //open the document
                document.Open();
                //add page to PdfCopy
                copy.AddPage(copy.GetImportedPage(reader, i));
                //close the document object
                document.Close();
            }

            return splitPages;
        }

        public static Dictionary<String, List<String>> ClassifyDocuments(List<String> pages)
        {

            List<String> native = new List<String>();
            List<String> scanned = new List<String>();
            Dictionary<String, List<String>> docTypes = new Dictionary<String, List<String>>();

            foreach (String page in pages)
            {
                // Get Text
                String docText = ExtractText(page);

                // Classify
                int wordCount = docText.Split(' ').Length;
                if(wordCount > 50)
                {
                    native.Add(page);
                }
                else
                {
                    scanned.Add(page);
                }
            }

            docTypes.Add("Native", native);
            docTypes.Add("Scanned", scanned);
            return docTypes;
        }

        public static string ExtractText(string filename)
        {
            StringBuilder text = new StringBuilder();

            if (File.Exists(filename))
            {
                PdfReader pdfReader = new PdfReader(filename);

                for (int page = 1;  page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);

                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    text.Append(currentText);
                }
                pdfReader.Close();
            }
            return text.ToString();
        }

        public static Dictionary<String, List<String>> ClassifyPODs(List<String> natives)
        {
            List<String> Pods = new List<String>();
            List<String> NonPods = new List<String>();
            Dictionary<String, List<String>> classification = new Dictionary<String, List<String>>();

            foreach (String page in natives)
            {
                String docText = ExtractText(page);
                bool isPod = IsPOD(docText);

                if(isPod == true)
                {
                    Pods.Add(page);
                }
                else
                {
                    NonPods.Add(page);
                }

            }

            classification.Add("Pods", Pods);
            classification.Add("NonPods", NonPods);

            return classification;
        }

        private static bool  IsPOD(String docText)
        {
            int successCount = 0;
            List<String> patterns = new List<String>(){
                @"\bName Of(?:\sthe\s|\s)Creditor\b",
                @"\bAddress Of(?:\sthe\s|\s)Creditor\b",
                @"\bTotal amount of(?:\sthe\s|\s)claim\b",
                @"\bSignature of(?:\sthe\s|\s)creditor\b"
            };

            foreach(String pattern in patterns)
            {
                if(Regex.Match(docText, pattern, RegexOptions.IgnoreCase).Success)
                {
                    successCount++;
                }
            }

            if(successCount >= 3)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

    }
}
