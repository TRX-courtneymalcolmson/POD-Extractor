using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace POD_Extractor
{
    class DocumentExtraction
    {
        public static List<Features> ExtractFeatures(List<String> spreadsheets)
        {
            List<Features> featureList = new List<Features>();
            foreach (String item in spreadsheets)
            {
                List<Features> features = Extract(item);
                foreach(Features _features_ in features)
                {
                    featureList.Add(_features_);
                }
            }
            return featureList;
        }

        private static List<Features> Extract(String path)
        {
            ExcelWorksheet ws =ExcelOperations.OpenExcelObject(path);

            List<Features> features = new List<Features>();

            List<String> CreditorReference = ExtractCreditorReference(ws);
            String CreditorName = ExtractCreditorName(ws);
            String CreditorPostcode = ExtractCreditorPostcode(ws);
            String ClientReference = ExtractClientReference(ws);
            String ClientName = ExtractClientName(ws);
            List<String> ClaimAmount = ExtractClaimAmount(ws);
    

            if (CreditorReference.Count == ClaimAmount.Count)
            {
                for(int i = 0; i < ClaimAmount.Count; i++)
                {
                    Features extractedFeatures = new Features();
                    extractedFeatures.CreditorReference = CreditorReference[i];
                    extractedFeatures.CreditorName = CreditorName;
                    extractedFeatures.CreditorPostcode = CreditorPostcode;
                    extractedFeatures.ClientReference = ClientReference;
                    extractedFeatures.ClientName = ClientName;
                    extractedFeatures.ClaimAmount = ClaimAmount[i];
                    extractedFeatures.PageFoundOn = Path.GetFileNameWithoutExtension(path);
                    extractedFeatures.Type = "PODDDDDD1234";
                    extractedFeatures.Valid = ValidateFeatures(extractedFeatures);
                    
                    features.Add(extractedFeatures);
                }
            }
            else
            {
                Features extractedFeatures = new Features();
                extractedFeatures.Valid = false;
                extractedFeatures.ErrorMessage = "Cannot match creditor references to claim amounts";
                features.Add(extractedFeatures);
            }

            return features;
        }

        
        private static List<String> ExtractCreditorReference(ExcelWorksheet ws)
        {
            List<String> matchedReferences = new List<String>();
           
                String value = ExcelOperations.SearchForPatterns(ws, Config.CreditorReferencePatterns);
                if (value == null)
                {
                    value = ExtractCreditorReferenceFromTable(ws) + ("TEST");
                }

            return matchedReferences;
        }

        private static String ExtractCreditorName(ExcelWorksheet ws)
        {
            try
            {
                String value = ExcelOperations.SearchForPatterns(ws, Config.CreditorNamePatterns);
                return value;
                //return CleanCreditorName(value);
            }
            catch
            {
                return null;
            }
        }
        private static String ExtractCreditorPostcode(ExcelWorksheet ws)
        {
            try
            {
                String value = ExcelOperations.SearchForPatterns(ws, Config.CreditorAddressPatterns);
                return CleanCreditorPostcode(value);
            }
            catch
            {
                return null;
            }
        }
        private static String ExtractClientReference(ExcelWorksheet ws)
        {
            try
            {
                String value = ExcelOperations.SearchForPatterns(ws, Config.ClientReferencePatterns);
                return CleanClientReference(value);
            }
            catch
            {
                return null;
            }
        }
        private static String ExtractClientName(ExcelWorksheet ws)
        {
            try
            {
                String value = ExcelOperations.SearchForPatterns(ws, Config.ClientNamePatterns);
                return CleanClientName(value);
            }
            catch
            {
                return null;
            }
        }
        private static List<String> ExtractClaimAmount(ExcelWorksheet ws)
        {
            List<String> matchedAmounts = new List<String>();
            try
            {
                String value = ExcelOperations.SearchForPatterns(ws, Config.ClaimAmountPatterns);
                value = RemoveMultipleClaimAmountsTotal(value);

                List<String> matches = SplitMultipleClaimAmounts(value);
                foreach(String match in matches)
                {
                    matchedAmounts.Add(CleanClaimAmount(match));
                }
                return matchedAmounts;
            }
            catch
            {
                return null;
            }
        }

        private static String ExtractCreditorReferenceFromTable(ExcelWorksheet ws)
        {
            String value = default;

            // Search Table for Account
            value = ExcelOperations.IrregularSpreadsheetSearch(ws,
                @"Account reference:",
                @"\s+([\w\/]+)");

            if (value != null)
            {
                return value;
            }

            // Search Table For Details
            value = ExcelOperations.IrregularSpreadsheetSearch(ws,
                @"(\b|^)Details of any document",
                @"([\s\w]*[\d\/\\]+[\w]*)");
            if (value != null) 
            { 
                return value; 
            }

            // Search Table For Creditor Name
            value = ExcelOperations.IrregularSpreadsheetSearch(ws,
                @"\bName Of(?:\sthe\s|\s)Creditor\b",
                @"([\s\w]*[\d\/\\]+[\w]*)");

            if (value != null)
            {
                return value + ("TEST");
            }

            return null;
        }

        private static string CleanCreditorReference(String value)
        {
            /// <summary>
            /// Method looks for common, mistakenly captured, components in References String and removes irrelevant components
            /// Irrelevant String components are matched with REGEX patterns and replaced bit "" values
            /// <\summary>
            /// 
            /// <param name="value"> Creditor Reference value to be cleaned </param>
            /// 
            /// <returns>
            /// Cleaned Creditor Reference
            /// </returns>

            if (value == null)
            {
                return null;
            }
            else
            {
                string pattern_1 = @"(?:(?:^|\b)Account reference: ((?:[A-z]*\d)+)";
                string pattern_2 = @"((?:[A-z]*\d)+)";
                Match matchVal = default;

                matchVal = Regex.Match(value, pattern_1, RegexOptions.IgnoreCase);
                if (matchVal.Success != false)
                {
                    return matchVal.Groups[matchVal.Groups.Count - 1].Value.Trim();
                }
                else
                {
                    matchVal = Regex.Match(value, pattern_2, RegexOptions.IgnoreCase);
                    if (matchVal.Success != false)
                    {
                        return matchVal.Groups[matchVal.Groups.Count - 1].Value.Trim();
                    }
                    else
                    {
                        return null;
                    }
                }
            }

        }

        private static string CleanClientReference(String value)
        {

            if (value == null)
            {
                return null;
            }
            else
            {
                string pattern_1 = @"(?:^|\b)(Your Ref[erence:]|Account reference)*([\s\d]*[\d\/\\]+[\w]*)";
                string pattern_2 = @"(:?Account reference:?)\s+(\d|\w{6,})*";
                Match matchVal = default;

                matchVal = Regex.Match(value, pattern_1, RegexOptions.IgnoreCase);
                if (matchVal.Success != false)
                {
                    return matchVal.Groups[matchVal.Groups.Count - 1].Value.Trim();
                }
                else
                {
                    matchVal = Regex.Match(value, pattern_2, RegexOptions.IgnoreCase);
                    if (matchVal.Success != false)
                    {
                        return matchVal.Groups[matchVal.Groups.Count - 1].Value.Trim();
                    }
                    else
                    {
                        return null;
                    }

                }

            }

        }
        private static string CleanClaimAmount(String value)
        {

            if (value == null)
            {
                return null;
            }
            else
            {
                //string pattern_1 = @"(?:(?:^|\b)£*)([\d,]*\.?\d*)";
                string pattern_1 = @"(?:^|\b)£*(\d+\.*\d*)";
                Match matchVal = default;

                matchVal = Regex.Match(value, pattern_1, RegexOptions.IgnoreCase);
                if (matchVal.Success != false)
                {
                    if (matchVal.Captures.Count == 1)
                    {
                        return matchVal.Groups[matchVal.Captures.Count - 1].Value.Trim().Replace("£","");
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }

            }

        }
        private static string CleanCreditorPostcode(String value)
        {

            if (value == null)
            {
                return null;
            }
            else
            {
                // Pattern found at: https://stackoverflow.com/questions/164979/regex-for-matching-uk-postcodes
                string pattern_1 = @"([Gg][Ii][Rr] 0[Aa]{2})|((([A-Za-z][0-9]{1,2})|(([A-Za-z][A-Ha-hJ-Yj-y][0-9]{1,2})|(([A-Za-z][0-9][A-Za-z])|([A-Za-z][A-Ha-hJ-Yj-y][0-9][A-Za-z]?))))\s?[0-9][A-Za-z]{2})";
                Match matchVal = default;

                matchVal = Regex.Match(value, pattern_1, RegexOptions.IgnoreCase);
                if (matchVal.Success != false)
                {
                    if (matchVal.Captures.Count == 1)
                    {
                        return matchVal.Groups[matchVal.Captures.Count - 1].Value.Trim();
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }

            }

        }
        public static string CleanClientName(String value)
        {
            if (value == null)
            {
                return null;
            }
            else
            {
                string name = default;
                string pattern_1 = @"(?:In The Matter Of:?|On behalf of:?|Re:|Client Details:?|Customers Name:?|NAME:?)";
                string pattern_2 = @"(?:under[ a]* Voluntary Arrangement [and]*|in[ a]* INDIVIDUAL VOLUNTARY ARRANGEMENT|[- ]* Proposed Individual Voluntary Arrangement)";
                string pattern_3 = @"[^a-z\s]+";

                name = Regex.Replace(value, pattern_1, "", RegexOptions.IgnoreCase);
                name = Regex.Replace(name, pattern_2, "", RegexOptions.IgnoreCase);
                name = Regex.Replace(name, pattern_3, "", RegexOptions.IgnoreCase);
                name = RemoveHonorifics(name);
                name = RemoveFalsePositives(name);
                return name.Trim();
            }

        }
        private static string RemoveHonorifics(String name)
        {
            string pattern = @"(MRS|MR|MASTER|MISS|MS|MX|DR)";
            return Regex.Replace(name, pattern, "", RegexOptions.IgnoreCase);
        }
        private static string RemoveFalsePositives(String name)
        {
            string pattern = @"(EE|Vanguard reference)";
            return Regex.Replace(name, pattern, "", RegexOptions.IgnoreCase);
        }

        private static bool ValidateFeatures(Features features)
        {
            if(features.CreditorReference == null && 
                features.CreditorName == null && 
                features.CreditorPostcode == null)
            {
                return false;
            }

            if(features.ClientName == null && features.ClientReference == null)
            {
                return false;
            }

            if(features.ClaimAmount == null)
            {
                return false;
            }

            return true;
        }

        private static List<String> SplitMultipleClaimAmounts(String value)
        {
            List<String> matchedAmounts = new List<String>();

            String pattern = @"(?:£|\b)\d+\.*\d*";
            MatchCollection matches = Regex.Matches(value, pattern);
            if(matches.Count > 0)
            {
                foreach(Match match in matches)
                {
                    if (match.Success != false)
                    {
                        matchedAmounts.Add( match.Groups[match.Groups.Count - 1].Value.Trim());
                    }
                }
            }

            return matchedAmounts;
        }
 
        private static String RemoveMultipleClaimAmountsTotal(String value)
        {
            String pattern = @"total\s*[\S\W]\s*£\d+\.*\d*";
            return Regex.Replace(value, pattern, "", RegexOptions.IgnoreCase);
        }

    }
}
