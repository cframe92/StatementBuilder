using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class MoneyPerksTransaction
    {
        public MoneyPerksTransaction()
        {
           
        }

        public DateTime Date
        {
            get;
            set;
        }

        public string Description
        {
            get;
            set;
        }

        public int Amount
        {
            get;
            set;
        }

        public int Balance
        {
            get;
            set;
        }
    }
}
