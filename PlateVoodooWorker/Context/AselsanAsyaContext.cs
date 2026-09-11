using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using PlateVoodooWorker.Model;

namespace PlateVoodooWorker.Context
{
    public class AselsanAsyaContext : DbContext
    {
        public DbSet<Aselsan> SAP_TRANSACTIONS_VIEW { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            string connectionString = @"User Id=GGMUSER;Password=resumgg;Data Source=172.46.10.11:1521/KMOASYA;";
            optionsBuilder.UseOracle(connectionString);
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);
            modelBuilder
                .Entity<Aselsan>()
                .HasNoKey()
                .ToView(null);
        }

    }
}
