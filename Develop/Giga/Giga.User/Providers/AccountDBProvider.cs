using System;
using System.Collections.Generic;
using System.Data.Entity.Migrations;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Giga.User.Providers
{
    /// <summary>
    /// Account provider that using DB to store account information
    /// </summary>
    public class AccountDBProvider : AccountProvider
    {
        private AccountDBContext _db = null;

        public override bool IsValid()
        {
            if (DB.Accounts.Where(a => a.ID == "Administrator").Count() < 1)
                return false;
            else
                return true;
        }

        /// <summary>
        /// Initialize Account database
        /// </summary>
        public override void Install()
        {
            DB.Accounts.AddOrUpdate(
                                    new Account
                                    {
                                        ID = "Administrator",
                                        Name = "Administrator",
                                        CreatedBy = "SYSTEM",
                                        CreatedTime = DateTime.Now,
                                        ModifiedBy = "SYSTEM",
                                        ModifiedTime = DateTime.Now,
                                        Password = Account.EncryptPassword("111111")
                                    });
        }

        public override void Uninstall(bool removeData)
        {
            Delete("Administrator");
        }

        public override void Initialize(Configuration.AccountProviderConfigurationElement cfg)
        {
            base.Initialize(cfg);
        }

        protected AccountDBContext DB
        {
            get
            {
                if (_db == null)
                {
                    _db = new AccountDBContext(ConnectStringName);
                }
                return _db;
            }
        }

        /// <summary>
        /// Create account
        /// </summary>
        /// <param name="info"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public override Account Create(Account info, string password)
        {
            //info.CreatedTime = DateTime.Now;
            //info.ModifiedTime = info.CreatedTime;
            DB.Accounts.Add(info);
            DB.SaveChanges();
            return info;
        }

        /// <summary>
        /// Delete account
        /// </summary>
        /// <param name="id"></param>
        public override void Delete(string id)
        {
            Account acc = (from a in DB.Accounts
                           where a.ID == id
                           select a).FirstOrDefault();
            if (acc != null)
            {
                DB.Accounts.Remove(acc);
                DB.SaveChanges();
            }
        }

        /// <summary>
        /// Update account
        /// </summary>
        /// <param name="account"></param>
        public override void Update(Account account)
        {
            Account exist = Query(a => a.ID == account.ID).FirstOrDefault();
            if (exist == null)
                throw new InvalidOperationException(String.Format("Account {0} not exist!", account.ID));
            exist.UpdateFrom(account);
            DB.SaveChanges();
        }

        /// <summary>
        /// Query account
        /// </summary>
        /// <param name="predicate">Condition expression</param>
        public override IQueryable<Account> Query(Expression<Func<Account, bool>> predicate)
        {
            return DB.Accounts.Where(predicate);
        }
    }
}
