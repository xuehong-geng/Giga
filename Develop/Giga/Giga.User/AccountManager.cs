using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
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

        private IAccountProvider _provider = null;
        /// <summary>
        /// Load account provider
        /// </summary>
        private void LoadProvider()
        {
            UserConfigurationSection sec = ConfigurationManager.GetSection("Giga.User") as UserConfigurationSection;
            if (sec == null)
            {
                LogManager.Error("AccountManager", "Cannot read configuration section Giga.User! No account provider available!");
            }
        }
    }
}
