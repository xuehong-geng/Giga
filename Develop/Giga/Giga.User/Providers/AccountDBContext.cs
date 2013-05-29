using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Text;
using System.Threading;
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

        public override int SaveChanges()
        {
            // Handle Modified & Created fields
            DateTime now = DateTime.Now;
            String curUser = Thread.CurrentPrincipal.Identity.Name;
            foreach (DbEntityEntry<Account> entry in ChangeTracker.Entries<Account>())
            {
                if (entry.State == System.Data.EntityState.Added)
                {
                    entry.Entity.CreatedBy = curUser;
                    entry.Entity.CreatedTime = now;
                    entry.Entity.ModifiedBy = curUser;
                    entry.Entity.ModifiedTime = now;
                }
                else if (entry.State == System.Data.EntityState.Modified)
                {
                    entry.Entity.ModifiedBy = curUser;
                    entry.Entity.ModifiedTime = now;
                }
            }
            return base.SaveChanges();
        }
    }
}
