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
        public double growth_gros1 { get; set; }
        public double growth_net1 { get; set; }
        public double share_gros1 { get; set; }
        public double share_net1 { get; set; }
        public double growth_gros2 { get; set; }
        public double growth_net2 { get; set; }
        public double share_gros2 { get; set; }
        public double share_net2 { get; set; }
        public string timestamp { get; set; }
    }

    public class ZuelligPharma_CalculatedDBContext : DbContext
    {
        public DbSet<ZuelligPharma_Calculated> ZuelligPharma_MATs { get; set; }
    }
}