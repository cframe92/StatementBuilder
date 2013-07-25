using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class Check
    {
        public Check(string checkNumber, decimal amount, DateTime date)
        {
            this.CheckNumber = checkNumber;
            this.Amount = amount;
            this.Date = date;
        }

        public string CheckNumber
        {
            get;
            set;

        }

        public decimal Amount
        {
            get;
            set;

        }

        public DateTime Date
        {
            get;
            set;
        }
    }
}
