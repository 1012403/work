using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace ZuelligPharma.Models
{
    public class ZuelligPharma_Frequency
    {
        public string adddt { get; set; }
        public string seqno { get; set; }
        public string freqno { get; set; }
        public int numofcust { get; set; }
        public double percentofcust { get; set; }
        public string timestamp { get; set; }
    }

    public class ZuelligPharma_FrequencyDBContext : DbContext
    {
        public DbSet<ZuelligPharma_Frequency> ZuelligPharma_Frequencys { get; set; }
    }
}