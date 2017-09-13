using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace UAStat
{

    public class UAStatContext : DbContext
    {
        public DbSet<UserAccount> Users { get; set; }
        public UAStatContext()
          : base()
        {

        }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {          
            optionsBuilder.UseNpgsql("User ID=тест;Password=тест;Server=тест;Port=тест;Database=тест;Pooling=true;");
        }
    }

}
