using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Patient_Master
{
    public class clsMap
    {
        public string Abstract { get; set; }
        public string Current_Assignee { get; set; }
        public string status { get; set; }
        public string Classfication { get; set; }
        public string Claim { get; set; }
        public string imgUrl { get; set; }
        public string Description { get; set; }
        public string anticipationExpiry { get; set; }
        public string title { get; set; }
    }
    public class patentCitation
    {
        public string publication_number { get; set; }
        public string priority_date { get; set; }
        public string publication_date { get; set; }
        public string assignee { get; set; }
        public string title { get; set; }
    }
    public class termsMap
    {
        public string patent_id { get; set; }
        public string terms { get; set; }
        public string cpc { get; set; }
    }
}
