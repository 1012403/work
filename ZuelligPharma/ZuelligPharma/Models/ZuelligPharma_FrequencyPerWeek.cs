using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace ZuelligPharma.Models
{
    public class ZuelligPharma_FrequencyPerWeek
    {
        public string adddt { get; set; }
        public string seqno { get; set; }
        public string week { get; set; }
        public int twice { get; set; }
        public int three { get; set; }
        public int more { get; set; }
        public string timestamp { get; set; }
    }

    public class ZuelligPharma_FrequencyPerWeekDBContext : DbContext
    {
        public DbSet<ZuelligPharma_FrequencyPerWeek> ZuelligPharma_FrequencyPerWeeks { get; set; }
    }
}