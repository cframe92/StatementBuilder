using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class Deposit : SortTableItem
    {
        public Deposit()
            : base()
        {

        }

        public Deposit(string description, decimal amount, DateTime date)
            : base (description, amount, date)
        {

        }
    }
}
