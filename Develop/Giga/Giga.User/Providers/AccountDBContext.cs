using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Giga.User.Providers
{
    /// <summary>
    /// DbContext for Account DB
    /// </summary>
    public class AccountDBContext : DbContext
    {
        public DbSet<Account> Accounts { get; set; }

        public AccountDBContext()
        {
        }
        public AccountDBContext(string ConnectStringName)
            : base(ConnectStringName)
        {
        }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            // Accounts
            modelBuilder.Entity<Account>().HasKey(a => new { a.ID })
                .ToTable("Account", "Giga");
            modelBuilder.Entity<Account>().Property(a => a.Password).HasMaxLength(256);
        }

        
    }
}
