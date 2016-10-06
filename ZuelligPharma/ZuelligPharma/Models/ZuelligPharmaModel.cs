using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ZuelligPharma.Models
{
    public class ZuelligPharmaModel
    {
        public List<ZuelligPharma_MAT> ZuelligPharma_MATs
        {
            get;
            set;
        }
        public List<ZuelligPharma_TopPRN> ZuelligPharma_TopPRNs
        {
            get;
            set;
        }
        public List<ZuelligPharma_ZPV> ZuelligPharma_ZPVs
        {
            get;
            set;
        }
        public List<ZuelligPharma_Calculated> ZuelligPharma_Calculateds
        {
            get;
            set;
        }
        public List<ZuelligPharma_Frequency> ZuelligPharma_Frequencys
        {
            get;
            set;
        }
        public List<ZuelligPharma_FrequencyPerWeek> ZuelligPharma_FrequencyPerWeeks
        {
            get;
            set;
        }
    }
}