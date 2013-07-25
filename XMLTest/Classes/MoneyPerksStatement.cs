using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class MoneyPerksStatement
    {
        public MoneyPerksStatement(string accountNumber)
        {
            AccountNumber = accountNumber;
            Transactions = new List<MoneyPerksTransaction>();
        }

        public string AccountNumber
        {
            get;
            set;
        }

        public List<MoneyPerksTransaction> Transactions
        {
            get;
            set;
        }

        public int BeginningBalance
        {
            get;
            set;
        }

        public int EndingBalance
        {
            get;
            set;
        }
    }
}
