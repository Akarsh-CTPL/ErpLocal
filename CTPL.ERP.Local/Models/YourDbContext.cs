using CTPL.ERP.Local.Data;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace CTPL.ERP.Local.Models
{
    public class YourDbContext : DbContext
    {
        public DbSet<Internal_Sims_Activations> InternalSimsActivationsModels { get; set; }

        public YourDbContext() : base("ERP_LocalEntities") 
        {
        }
    }
}