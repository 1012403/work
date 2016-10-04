using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace ZuelligPharma.Models
{
    public class ZuelligPharma_TopPRN
    {
        public string adddt { get; set; }
        public string seqno { get; set; }
        public string prnkey { get; set; }
        public string monthfr { get; set; }
        public string monthto { get; set; }
        public double sale_monthfr { get; set; }
        public double sale_monthto { get; set; }
        public double month_growth { get; set; }
        public double month_share { get; set; }
        public string yearfr { get; set; }
        public string yearto { get; set; }
        public double sale_yearfr { get; set; }
        public double sale_yearto { get; set; }
        public double year_growth { get; set; }
        public double year_share { get; set; }
        public string timestamp { get; set; }
    }

    public class ZuelligPharma_TopPRNDBContext : DbContext
    {
        public DbSet<ZuelligPharma_TopPRN> ZuelligPharma_TopPRNs { get; set; }
    }
}