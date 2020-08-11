using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace POD_Extractor
{
    class Features
    {
        public String CreditorReference { get; set; }
        public String CreditorName { get; set; }
        public String CreditorPostcode { get; set; }
        public String ClientReference { get; set; }
        public String ClientName { get; set; }
        public String ClaimAmount { get; set; }

        public String PageFoundOn { get; set; }
        public String Type { get; set; }
        public bool Valid { get; set; }

        public String ErrorMessage { get; set; }

    }
}
