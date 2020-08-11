using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POD_Extractor
{
    class Reporting
    {
        public String Filename { get; set; }
        public List<Features> ExtractedFeatures { get; set; }
        public List<Features> Scans { get; set; }
        public List<Features> NonPods { get; set; }


        public static List<Features> Package(List<String> iterable, String type)
        {
            List<Features> payload = new List<Features>();
            foreach(String page in iterable)
            {
                Features features = new Features();
                features.Type = type;
                features.PageFoundOn = Path.GetFileNameWithoutExtension(page);
                payload.Add(features);
                
            }
            return payload;
        }

        public static String SerializeReport(Reporting report)
        {
            return JsonConvert.SerializeObject(report);
        }
    }
}
