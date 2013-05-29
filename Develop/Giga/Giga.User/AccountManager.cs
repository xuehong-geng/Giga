using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Giga.Log;
using Giga.User.Configuration;

namespace Giga.User
{
    /// <summary>
    /// Account management controller
    /// </summary>
    public class AccountManager
    {
        /// <summary>
        /// Singletone controller
        /// </summary>
        private static AccountManager _instance = new AccountManager();
        private static AccountManager GetInstance() { return _instance; }

        private AccountManager()
        {
            LoadProvider();
        }

        private AccountProvider _provider = null;
        /// <summary>
        /// Load account provider
        /// </summary>
        private void LoadProvider()
        {
            UserConfigurationSection sec = ConfigurationManager.GetSection("Giga.User") as UserConfigurationSection;
            if (sec == null)
            {
                LogManager.Error("Cannot read configuration section Giga.User! No account provider available!");
                return;
            }
            AccountProviderConfigurationElement provElem = sec.AccountProviders.Get(sec.AccountProvider);
            if (provElem == null)
            {
                LogManager.Error("Cannot find configuration for Account Provider {0}!", sec.AccountProvider);
                return;
            }
            // Create provider instance
            try
            {
                Type type = Type.GetType(provElem.Type);
                _provider = type.Assembly.CreateInstance(type.FullName) as AccountProvider;
            }
            catch (Exception err)
            {
                LogManager.Error(err, "Create Account Provider {0} failed!", provElem.Type);
                return;
            }
            // Initialize the provider
            try
            {
                _provider.Initialize(provElem);
            }
            catch (Exception err)
            {
                LogManager.Error(err, "Initialize account provider {0} failed!", provElem.Name);
                _provider = null;
                return;
            }
            // Check if environment is ok
            if (!_provider.IsValid())
            {   // The dependencies of provider is not valid, re-install it to fix these problems
                try
                {
                    _provider.Install();
                }
                catch (Exception err)
                {
                    LogManager.Error(err, "Cannot install account provider {0}!", provElem.Name);
                    _provider = null;
                    return;
                }
            }
        }

        /// <summary>
        /// Get current account provider
        /// </summary>
        /// <returns></returns>
        private IAccountProvider GetProvider()
        {
            if (_provider == null)
            {
                throw new InvalidOperationException("No Account Provider exists! Please check configuration.");
            }
            return _provider;
        }

        /// <summary>
        /// Create new user account
        /// </summary>
        /// <param name="info">Information of new account</param>
        /// <param name="password">Password of new account</param>
        /// <returns>Created account</returns>
        public static Account CreateAccount(Account info, String password)
        {
            return GetInstance().GetProvider().Create(info, password);
        }

        /// <summary>
        /// Delete user account
        /// </summary>
        /// <param name="account">ID of user account to be deleted</param>
        public static void DeleteAccount(String id)
        {
            GetInstance().GetProvider().Delete(id);
        }

        /// <summary>
        /// Update account 
        /// </summary>
        /// <param name="account">Account info</param>
        public static void UpdateAccount(Account account)
        {
            GetInstance().GetProvider().Update(account);
        }

        /// <summary>
        /// Query account
        /// </summary>
        /// <param name="predicate">Query condition</param>
        /// <returns></returns>
        public static IQueryable<Account> QueryAccount(Expression<Func<Account, bool>> predicate)
        {
            return GetInstance().GetProvider().Query(predicate);
        }

        /// <summary>
        /// Get account by ID
        /// </summary>
        /// <param name="id">ID of account</param>
        /// <returns></returns>
        public static Account GetAccount(String id)
        {
            IQueryable<Account> rs = QueryAccount(a => a.ID == id);
            return rs.FirstOrDefault();
        }
    }
}
