using Microsoft.EntityFrameworkCore;
using PlateVoodooWorker.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PlateVoodooWorker.Context
{
    public class PlateVoodooContext : DbContext
    {
        public DbSet<Vehicle> Vehicles { get; set; }


        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            string connectionString = "Host=10.75.25.57:5432;Username=postgres; Password=XZf3ccbJnEF4W15VMj6w; Database=Kmo.PlateVoodoo";
            optionsBuilder.UseNpgsql(connectionString);

        }
    }
}
