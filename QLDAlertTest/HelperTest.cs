using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using QLDAlert;

namespace QLDAlertTest {
    [TestClass]
    public class HelperTest {
        private DateTime today = System.DateTime.Today.Date + new TimeSpan(0, 0, 0);

        #region Tests on isGoingToExpire()

        [TestMethod]
        public void isGoingToExpireTestedByYesterday() {
            Helper helper = new Helper();
            var yesterday = today.AddDays(-1);
            Assert.IsTrue(helper.isGoingToExpire(yesterday));
        }

        [TestMethod]
        public void isGoingToExpireTestedByDateInThreeWeeks() {
            Helper helper = new Helper();
            var threeWeeksLater = today.AddDays(21);
            Assert.IsTrue(helper.isGoingToExpire(threeWeeksLater));
        }

        [TestMethod]
        public void isGoingToExpireTestedByDateInSixWeeks() {
            Helper helper = new Helper();
            var sixWeeksLater = today.AddDays(42);
            Assert.IsTrue(helper.isGoingToExpire(sixWeeksLater));
        }

        [TestMethod]
        public void isGoingToExpireTestedByDateInMoreThanSixWeeks() {
            Helper helper = new Helper();
            var sixWeeksAndOneDayLater = today.AddDays(43);
            Assert.IsFalse(helper.isGoingToExpire(sixWeeksAndOneDayLater));
        }

        #endregion

        #region Tests on isMemberActive()

        [TestMethod]
        public void isMemberActiveTestedByOneYearAgo() {
            Helper helper = new Helper();
            var oneYearAgo = today.AddYears(-1);
            Assert.IsTrue(helper.isMemberActive(oneYearAgo));
        }

        [TestMethod]
        public void isMemberActiveTestedByAlmostTwoYearsAgo() {
            Helper helper = new Helper();
            var almostTwoYearsAgo = today.AddYears(-2);
            almostTwoYearsAgo = almostTwoYearsAgo.AddDays(1);

            Assert.IsTrue(helper.isMemberActive(almostTwoYearsAgo));
        }

        [TestMethod]
        public void isMemberActiveTestedByTwoYearsAgo() {
            Helper helper = new Helper();
            var twoYearAgo = today.AddYears(-2);
            Assert.IsFalse(helper.isMemberActive(twoYearAgo));
        }

        [TestMethod]
        public void isMemberActiveTestedByMoreTwoYearsAgo() {
            Helper helper = new Helper();
            var twoYearAgo = today.AddYears(-2);
            twoYearAgo = twoYearAgo.AddDays(-1);
            Assert.IsFalse(helper.isMemberActive(twoYearAgo));
        }

        #endregion

        #region Tests on isDateEmpty()

        [TestMethod]
        public void isDateEmptyTestedByEmptyDate() {
            Helper helper = new Helper();
            var dateTime = new DateTime();
            Assert.IsTrue(helper.isDateEmpty(dateTime));
        }

        [TestMethod]
        public void isDateEmptyTestedByValidDate() {
            Helper helper = new Helper();
            var today = DateTime.Today;
            Assert.IsFalse(helper.isDateEmpty(today));
        }

        #endregion
    }
}
