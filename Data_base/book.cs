using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Data_base
{
    class book
    {
        public int id { get; set; }
        public string name { get; set; }
        public string authorFirstName { get; set; }
        public string authroLastName { get; set; }
        public string press { get; set; }
        public double price { get; set; }
        public book(int id, string name, string authorFirstName,
            string authroLastName, string press, double price)
        {
            this.name = name;
            this.authorFirstName = authorFirstName;
            this.authroLastName = authroLastName;
            this.press = press;
            this.price = price;
        }
    }
}
