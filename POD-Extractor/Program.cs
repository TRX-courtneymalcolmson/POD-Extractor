using System;
using System.Collections.Generic;
using System.IO;

namespace POD_Extractor
{
    class Program
    {

        static void Main(string[] args)
        {
            //String TEMP_PDF = @"C:\Users\Laure\Documents\Vanguard\pdfs\SafetyNet\101048928.pdf";
            //List<String> ocrDocs = new List<String>() { @"C:\ProgramData\POD-Extractor\TEMP_SPREADSHEET\1.xlsx" }; // For testing only
            String jsonReport = default;
            try
            {
                String TEMP_PDF = @"C:\Users\Laure\Documents\Vanguard\pdfs\SafetyNet\101048928.pdf";

                // Prepare for work
                TEMP_PDF = Admin.ProgramSetup(TEMP_PDF);

                // Split pdf pages into seperate pages
                List<String> pdfPages = PdfOperations.SplitPdfPages(TEMP_PDF, Config.TEMP_PAGESPLIT);

                // Classify PDF documents as either native or Scans
                // Do something with non-natives
                Dictionary<String, List<String>> docTypes = PdfOperations.ClassifyDocuments(pdfPages);

                // Determin which pages are PODs
                List<String> natives = new List<String>();
                docTypes.TryGetValue("Native", out natives);
                var podClassifications = PdfOperations.ClassifyPODs(natives);

                // Apply OCR to POD Documents
                List<String> pods = new List<String>();
                podClassifications.TryGetValue("Pods", out pods);
                var ocrDocs = AbbyyOCR.PdfToXlsx(pods);


                // ---- Package Results into Features Objects ---- 
                // PODS
                var nativeFeatures = DocumentExtraction.ExtractFeatures(ocrDocs);

                // SCANS
                List<String> _scanned_ = new List<String>();
                docTypes.TryGetValue("Scanned", out _scanned_);
                var scans = Reporting.Package(_scanned_, "SCANNED");

                // Non-Pods
                List<String> _nonPods_ = new List<String>();
                podClassifications.TryGetValue("NonPods", out _nonPods_);
                var nonPods = Reporting.Package(_nonPods_, "NONPOD");

                // Json
                Reporting finalReport = new Reporting();
                finalReport.Filename = Path.GetFileName(TEMP_PDF);
                finalReport.ExtractedFeatures = nativeFeatures;
                finalReport.Scans = scans;
                finalReport.NonPods = nonPods;

                jsonReport = Reporting.SerializeReport(finalReport);

            }
            catch (Exception ex)
            {
                jsonReport = "0";
            }
            finally
            {
                File.WriteAllText(Config.TEMP_OUTPUT, jsonReport, System.Text.Encoding.UTF8);
            }

        }
    }
}
