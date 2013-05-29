using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace Giga.User
{
    /// <summary>
    /// Interface of provider that provide functionalities for account management
    /// </summary>
    public interface IAccountProvider
    {
        /// <summary>
        /// Create new user account
        /// </summary>
        /// <param name="info">Information of new account</param>
        /// <param name="password">Password of new account</param>
        /// <returns>Account created</returns>
        Account Create(Account info, String password);

        /// <summary>
        /// Delete user account
        /// </summary>
        /// <param name="id">Id of account to be deleted</param>
        void Delete(string id);

        /// <summary>
        /// Update user account
        /// </summary>
        /// <param name="account"></param>
        void Update(Account account);

        /// <summary>
        /// Query account
        /// </summary>
        /// <param name="predicate">Condition expression</param>
        IQueryable<Account> Query(Expression<Func<Account, bool>> predicate);
    }
}
