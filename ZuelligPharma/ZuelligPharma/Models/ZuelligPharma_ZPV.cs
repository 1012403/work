using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace ZuelligPharma.Models
{
    public class ZuelligPharma_ZPV
    {
        private string adddt { get; set; }
        private string seqno { get; set; }
        private string ytd { get; set; }
        private decimal sanofi { get; set; }
        private decimal gsk { get; set; }
        private decimal msd { get; set; }
        private decimal az { get; set; }
        private decimal pfitzer { get; set; }
        private decimal bayer { get; set; }
        private decimal topprn { get; set; }
        private string typevalue { get; set; }
        private string timestamp { get; set; }
    }

    public class ZuelligPharma_ZPVDBContext : DbContext
    {
        public DbSet<ZuelligPharma_ZPV> ZuelligPharma_ZPVs { get; set; }
    }
}