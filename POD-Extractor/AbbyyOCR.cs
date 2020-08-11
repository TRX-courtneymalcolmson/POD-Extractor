using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POD_Extractor
{
    class AbbyyOCR
    {
        public POD_Extractor.EngineLoader engineLoader = null;

        public static List<String> PdfToXlsx(List<String> pods)
        {
            List<String> abbyyPaths = new List<String>();

            foreach(String pod in pods)
            {
                // Apply Abbyy
                String xlsxPath = PdfToXlsx(pod);
                // Return Filepaths
                abbyyPaths.Add(xlsxPath);
            }

            return abbyyPaths;
        }

        private static String PdfToXlsx(string path)
        {
            String xlsxPath = default;

            // Load Engine //
            EngineLoader engineLoader = new EngineLoader();
            engineLoader.Engine.LoadPredefinedProfile("DocumentConversion_Accuracy");

            // Process Engine //
            xlsxPath = ProcessEngine(engineLoader, path);

            // Dispose Engine //
            engineLoader.Dispose();


            return xlsxPath;
        }

        public static String ProcessEngine(EngineLoader engineLoader, string path)
        {
            // Assign/Create paths needed for engine
            string outPath = Path.Combine(Config.TEMP_SPREADSHEET, Path.GetFileNameWithoutExtension(path)) + ".xlsx";

            // Create document
            FREngine.FRDocument document = engineLoader.Engine.CreateFRDocument();

            try
            {
                // Add image file to document
                document.AddImageFile(path, null, null);

                // Recognize document
                document.Process(null);

                //// Save results to pdf using 'balanced' scenario
                FREngine.PDFExportParams pdfParams = engineLoader.Engine.CreatePDFExportParams();
                pdfParams.Scenario = FREngine.PDFExportScenarioEnum.PES_Balanced;


                FREngine.XLExportParams XLparams = engineLoader.Engine.CreateXLExportParams();
                XLparams.OnePagePerWorksheet = true;

                document.Export(outPath, FREngine.FileExportFormatEnum.FEF_XLSX, XLparams);

            }
            catch (Exception error)
            {
                Console.WriteLine("ERROR PROCESSING ABBYY\n\n{0}", error.Message);
            }
            finally
            {
                // Close document
                document.Close();
            }
            return outPath;
        }

    }
}
