using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace XMLTest.Classes
{
    public class Configuration
    {
        public static string GetStatementsOutputPath()
        {
            return ConfigurationManager.AppSettings["StatementsOutputPath"].ToString();
        }

        public static string GetSymitarStatementDataFilePath()
        {
            return ConfigurationManager.AppSettings["SymitarStatementDataFilePath"].ToString();
        }

        public static string GetMoneyPerksFilePath()
        {
            return ConfigurationManager.AppSettings["MoneyPerksFilePath"].ToString();
        }

        public static string GetErrorLogOutputPath()
        {
            return ConfigurationManager.AppSettings["ErrorLogOutputPath"].ToString();
        }

        public static string GetStatementTemplateFirstPageFilePath()
        {
            return ConfigurationManager.AppSettings["StatementTemplateFirstPageFilePath"].ToString();
        }

        public static string GetStatementTemplateFilePath()
        {
            return ConfigurationManager.AppSettings["StatementTemplateFilePath"].ToString();
        }

        public static string GetStatementDisclosuresTemplateFilePath()
        {
            return ConfigurationManager.AppSettings["StatementDisclosuresTemplateFilePath"].ToString();
        }
    }
}
