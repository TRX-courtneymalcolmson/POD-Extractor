using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POD_Extractor
{
    class Admin
    {
        // This class is prepares the script for extraction and also
        // cleans-up after script is comeplete

        public static String ProgramSetup(String pdfPath)
        {
            ValidateInput(pdfPath);

            // Create Folders
            CreateDirectory(Config.PROGRAM_DATA_FOLDER);
            CreateDirectory(Config.TEMP_PAGESPLIT);
            CreateDirectory(Config.TEMP_SPREADSHEET);
            CreateDirectory(Config.TEMP_WORDTOPDF);

            // Delete any data in TEMP folders
            ClearTempDirectories(Config.TEMP_PAGESPLIT);
            ClearTempDirectories(Config.TEMP_SPREADSHEET);
            ClearTempDirectories(Config.TEMP_WORDTOPDF);

            // Delete any previous results
            if (File.Exists(Config.TEMP_OUTPUT))
            {
                File.Delete(Config.TEMP_OUTPUT);
            }

            // Convert Word Doc to PDF
            if (Path.GetExtension(pdfPath).ToUpper().Contains(".DOC")){
                return WordToPdf(pdfPath);
            }
            else
            {
                return pdfPath;
            }
        }

        private static void ValidateInput(String pdfPath)
        {
            if (File.Exists(pdfPath) == false)
            {
                throw new Exception("Supplied Path Cannot Be Found - Please Check File Exists");
            }

            if (!Path.GetExtension(pdfPath).ToUpper().Contains(".PDF") && !Path.GetExtension(pdfPath).ToUpper().Contains(".DOC"))
            {
                throw new Exception("Supplied Path Must be PDF - Please Check File Type");
            }
        }

        private static void CreateDirectory(String dirPath)
        {
            // Create directory if it doesn't exist
            if(Directory.Exists(dirPath) == false)
            {
                Directory.CreateDirectory(dirPath);
            }
        }

        private static void ClearTempDirectories(String dirPath)
        {
            // Delete any files from inputed directory
            String[] files = Directory.GetFiles(dirPath);
            if(files.Length != 0)
            {
                foreach (String file in files)
                {
                    File.Delete(file);
                }
            }
        }

        public static String WordToPdf(String filePath)
        {
            String savePath = Path.Combine(Config.TEMP_WORDTOPDF, "converted.pdf");
            Document doc = new Document(filePath);
            doc.Save(savePath, SaveFormat.Pdf);
            if(File.Exists(savePath) == true)
            {
                return savePath;
            }
            else
            {
                throw new Exception("Cannot Convert Document to PDF");
            }

        }
    }
}
