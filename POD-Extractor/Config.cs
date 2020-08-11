using System;
using System.Collections.Generic;
using System.IO;

namespace POD_Extractor
{
    class Config
    {
        // Program Folder Structure 
        public static String PROGRAM_DATA_FOLDER = @"C:\ProgramData\POD-Extractor";
        public static String TEMP_PAGESPLIT = Path.Combine(PROGRAM_DATA_FOLDER, "TEMP_PAGESPLIT");
        public static String TEMP_SPREADSHEET = Path.Combine(PROGRAM_DATA_FOLDER, "TEMP_SPREADSHEET");
        public static String TEMP_WORDTOPDF = Path.Combine(PROGRAM_DATA_FOLDER, "TEMP_WORDTOPDF");
        public static String TEMP_OUTPUT = PROGRAM_DATA_FOLDER + @"\output.json";


        // REGEX Patterns for Feature Extraction
        public static List<Tuple<string, int, int>> CreditorNamePatterns = new List<Tuple<string, int, int>>
        {
            new Tuple<string, int, int>(@"\bName Of(?:\sthe\s|\s)Creditor\b", 0, 1),
        };
        public static List<Tuple<string, int, int>> CreditorReferencePatterns = new List<Tuple<string, int, int>>
        {        
            new Tuple<string, int, int>(@"Account reference", 0, 1),           
        };
        public static List<Tuple<string, int, int>> CreditorAddressPatterns = new List<Tuple<string, int, int>>
        {
            new Tuple<string, int, int>(@"\bAddress Of(?:\sthe\s|\s)Creditor\b", 0, 1),
        };

        public static List<Tuple<string, int, int>> ClientNamePatterns = new List<Tuple<string, int, int>>
        {
            new Tuple<string, int, int>(@"(\b|^)Re:", 0, 0), // keep this as first in list always
            new Tuple<string, int, int>(@"Client Details", 0, 0),
            new Tuple<string, int, int>(@"Name in BLOCK LETTERS", 0, 1),
            new Tuple<string, int, int>(@"(^|\b|.)Name(?:(?!\bcreditor\b).)*$", 0, 0),
            new Tuple<string, int, int>(@"\w+ Individual Voluntary Arrangement", 0, 0),
            new Tuple<string, int, int>(@"In The Matter [of]*", 0, 0),

        };
        public static List<Tuple<string, int, int>> ClientReferencePatterns = new List<Tuple<string, int, int>>
        {
            new Tuple<string, int, int>(@"(\b|^)IP Case Ref", 0, 0),
            new Tuple<string, int, int>(@"(\b|^)Your Ref", 0, 0),
            new Tuple<string, int, int>(@"(\b|^)Your Reference", 0, 0),
        };

        public static List<Tuple<string, int, int>> ClaimAmountPatterns = new List<Tuple<string, int, int>>
        {
            new Tuple<string, int, int>(@"Total amount of claim", 0, 1),

        };

    }
}
