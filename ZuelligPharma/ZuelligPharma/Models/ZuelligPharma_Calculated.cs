using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace ZuelligPharma.Models
{
    public class ZuelligPharma_Calculated
    {
        public string adddt { get; set; }
        public string seqno { get; set; }
        public string area { get; set; }
        public string monthfr { get; set; }
        public string monthto { get; set; }
        public double sale_gros_monthfr { get; set; }
        public double sale_gros_monthto { get; set; }
        public double share_gros_month { get; set; }
        public double growth_gros_month { get; set; }
        public double sale_net_monthfr { get; set; }
        public double sale_net_monthto { get; set; }
        public double share_net_month { get; set; }
        public double growth_net_month { get; set; }
        public string ytdfr { get; set; }
        public string ytdto { get; set; }
        public double sale_gros_ytdfr { get; set; }
        public double sale_gros_ytdto { get; set; }
        public double share_gros_ytd { get; set; }
        public double growth_gros_ytd { get; set; }
        public double sale_net_ytdfr { get; set; }
        public double sale_net_ytdto { get; set; }
        public double share_net_ytd { get; set; }
        public double growth_net_ytd { get; set; }
        public string timestamp { get; set; }
    }

    public class ZuelligPharma_CalculatedDBContext : DbContext
    {
        public DbSet<ZuelligPharma_Calculated> ZuelligPharma_Calculateds { get; set; }
    }
}