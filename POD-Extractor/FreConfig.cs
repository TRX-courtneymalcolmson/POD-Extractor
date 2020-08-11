using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POD_Extractor
{
    class FreConfig
    {
        // Folder with FRE dll
        public static String GetDllFolder()
        {
            
            if (is64BitConfiguration())
            {
                return @"C:\Program Files\ABBYY SDK\12\FineReader Engine\Bin64";
            }
            else
            {
                // If statement enters this 'else' section, it is trying to load a 32bit dll and will fail.
                // Go to Project > Project Properties > Build - and set to target x64 not x32
                return @"C:\Program Files\ABBYY SDK\12\FineReader Engine\Bin64";
            }
        }

        // Return customer project id for FRE
        public static String GetCustomerProjectId()
        {
            return "2vpmByZt59fCqBqjJGPG";
        }

        // Return path to license file
        public static String GetLicensePath()
        {
            return "";
        }

        // Return license password
        public static String GetLicensePassword()
        {
            return "";
        }

        // Return full path to Samples directory
        public static String GetSamplesFolder()
        {
            return "C:\\ProgramData\\ABBYY\\SDK\\12\\FineReader Engine\\Samples";
        }

        // Determines whether the current configuration is a 64-bit configuration
        private static bool is64BitConfiguration()
        {
            return System.IntPtr.Size == 8;
        }
    }
}
