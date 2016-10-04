using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;

namespace ZuelligPharma.Models
{
    public class ZuelligPharma_MAT
    {
        public string adddt { get; set; }
        public string seqno { get; set; }
        public string date { get; set; }
        public double gros { get; set; }
        public double net { get; set; }
        public double sale { get; set; }
        public string timestamp { get; set; }
    }

    public class ZuelligPharma_MATDBContext : DbContext
    {
        public DbSet<ZuelligPharma_MAT> ZuelligPharma_MATs { get; set; }
    }
}