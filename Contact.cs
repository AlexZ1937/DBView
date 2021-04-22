using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DBView
{
    class Contact
    {
        string contactname;
        string contactnumber;
        public Contact(string name, string number)
        {
            this.contactname = name;
            this.contactnumber = number;
        }
        public string GetName()
        {
            return contactname;
        }

        public string GetNumber()
        {
            return contactnumber;
        }

    }
}
