using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Giga.User;

namespace Giga.Test.User
{
    [TestClass]
    public class AccountTest
    {
        [TestMethod]
        public void TestCreateAccount()
        {
            Account act = new Account();
            act.ID = Guid.NewGuid().ToString();
            act.Name = "Tester";
            AccountManager.CreateAccount(act, "123456");
            AccountManager.DeleteAccount(act.ID);
        }
    }
}
