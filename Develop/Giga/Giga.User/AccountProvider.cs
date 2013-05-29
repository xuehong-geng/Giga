using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using Giga.User.Configuration;

namespace Giga.User
{
    /// <summary>
    /// Base class of account provider
    /// </summary>
    public abstract class AccountProvider : IAccountProvider
    {
        /// <summary>
        /// Name of provider
        /// </summary>
        public String Name { get; set; }
        /// <summary>
        /// Name of connect string
        /// </summary>
        public String ConnectStringName { get; set; }

        /// <summary>
        /// Initialize the account provider
        /// </summary>
        /// <param name="cfg">Configuration element</param>
        public virtual void Initialize(AccountProviderConfigurationElement cfg)
        {
            Name = cfg.Name;
            ConnectStringName = cfg.ConnectStringName;
        }

        /// <summary>
        /// Check if environment is valid and the provider is ready to use
        /// </summary>
        /// <returns></returns>
        public abstract bool IsValid();

        /// <summary>
        /// Install and setup environment for provider.
        /// </summary>
        public abstract void Install();

        /// <summary>
        /// Uninstall environment of provider
        /// </summary>
        /// <param name="removeData">Whether remove all data</param>
        public abstract void Uninstall(bool removeData);

        public abstract Account Create(Account info, string password);

        public abstract void Delete(string id);

        public abstract void Update(Account account);

        public abstract IQueryable<Account> Query(Expression<Func<Account, bool>> predicate);
    }
}
