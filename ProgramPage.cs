using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Entity;

 namespace Selenium.NetCore.Test
{
    public class ProgramPage : DbContext
    {
        public ProgramPage() : base("Name=PomEntities")
        {
          this.Database.Connection.ConnectionString = String.Format("Data Source={0}", ".\\WwsPomProject.db");  
        }


        public DbSet<Used> ItemList { get; set; }
        
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Configurations.Add(new Used());
            base.OnModelCreating(modelBuilder);
        }
    }
}
