using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QLDAlert {
    class Member {
        public string name { get; set; }
        public string email { get; set; }
        public int postcode { get; set; }
        public DateTime expDateBlueCard { get; set; }
        public string medProfession { get; set; }
        public DateTime expDateMedPro { get; set; }
        public DateTime expDatePoliceCheck { get; set; }
        public DateTime lastActivity { get; set; }
        public string contactId { get; set; }

        public override string ToString() {
            string memberString = "";

            memberString += "Name: " + name + "\n";
            memberString += "Email: " + email + "\n";
            memberString += "Postcode: " + postcode + "\n";
            memberString += "Expiry Date - Blue Card: " + expDateBlueCard.ToString("dd/MM/yyyy") + "\n";
            memberString += "Medical Profession: " + medProfession + "\n";
            memberString +=
                "Expiry Date - Medical Profession: " + expDateMedPro.ToString("dd/MM/yyyy") + "\n";
            memberString +=
                "Expiry Date - Police Check: " + expDatePoliceCheck.ToString("dd/MM/yyyy") + "\n";
            memberString +=
                "Date of Last Activity: " + lastActivity.ToString("dd/MM/yyyy") + "\n";
            memberString +=
                "Contact ID: " + contactId + "\n";

            return memberString;
        }
    }
}
