using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLTest.Classes
{
    public class StatementFileField
    {
        public StatementFileField(string field)
        {
            Type = string.Empty;
            Data = string.Empty;

            if (field.Length == STATEMENT_FILE_RECORD_FIELD_TYPE_LENGTH)
            {
                Type = field;
            }
            else if (field.Length > STATEMENT_FILE_RECORD_FIELD_TYPE_LENGTH)
            {
                Type = field.Substring(0, STATEMENT_FILE_RECORD_FIELD_TYPE_LENGTH);
                Data = field.Substring(STATEMENT_FILE_RECORD_FIELD_TYPE_LENGTH);
            }
        }

        public string Type
        {
            get;
            set;
        }

        public string Data
        {
            get;
            set;
        }

        public const int STATEMENT_FILE_RECORD_FIELD_TYPE_LENGTH = 2;
    }
}
