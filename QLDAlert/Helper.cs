using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QLDAlert {
    public class Helper {
        private const string INPUT_DATE_FORMAT = "M/d/yyyy";
        private const string OUTPUT_DATE_FORMAT = "dd/MM/yyyy";
        private DateTime today;

        public Helper() {
            today = DateTime.Today;
            today = today.Date + new TimeSpan(0, 0, 0);
        }

        /// <summary>
        /// Checks if the expiry date is due in 6 weeks.
        /// Returns true if the date is within the next 6 weeks.
        /// </summary>
        /// <param name="dt"></param>
        /// <returns>Returns boolean</returns>
        public bool isGoingToExpire(DateTime dt) {
            TimeSpan diff;
            TimeSpan sixWeeks = new TimeSpan(42, 0, 0, 0);

            diff = dt - DateTime.Today;

            return diff <= sixWeeks;
        }

        /// <summary>
        /// Checks if date of last activity is earlier than 2 years ago.
        /// </summary>
        /// <param name="dt"></param>
        /// <returns>Retruns boolean</returns>
        public bool isMemberActive(DateTime dt) {
            TimeSpan diff;
            TimeSpan twoYears = today.AddYears(2) - today;

            var _dt = dt.Date + new TimeSpan(0, 0, 0);

            diff = today - _dt;

            return twoYears >= diff;
        }

        /// <summary>
        /// Returns true if the datetime input has no set value.
        /// </summary>
        /// <param name="dt"></param>
        /// <returns>Returns boolean</returns>
        public bool isDateEmpty(DateTime dt) {
            DateTime emptyDate = DateTime.ParseExact("1/1/0001", INPUT_DATE_FORMAT, null);

            return DateTime.Compare(emptyDate, dt) == 0;
        }
    }
}
