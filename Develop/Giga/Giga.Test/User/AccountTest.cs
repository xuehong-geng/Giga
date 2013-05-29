using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Giga.User;
using System.Threading;
using System.Linq;

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
            Account r = AccountManager.CreateAccount(act, "123456");
            AccountManager.DeleteAccount(r.ID);
        }

        [TestMethod]
        public void TestModifyAccount()
        {
            Account act = new Account();
            act.ID = Guid.NewGuid().ToString();
            act.Name = "Tester";
            Account r = AccountManager.CreateAccount(act, "123456");
            DateTime now = DateTime.Now;
            String user = Thread.CurrentPrincipal.Identity.Name;
            r.Email = "tester@gmail.com";
            r.MobilePhone = "13012938823";
            AccountManager.UpdateAccount(r);
            Account r1 = AccountManager.QueryAccount(a => a.ID == r.ID).FirstOrDefault();
            Assert.IsNotNull(r1);
            Assert.IsTrue(r.Email == r1.Email);
            Assert.IsTrue(r.MobilePhone == r1.MobilePhone);
            Assert.IsTrue(r.ModifiedBy == user);
            Assert.IsTrue(r.ModifiedTime >= now);
            AccountManager.DeleteAccount(r.ID);
        }
    }
}
