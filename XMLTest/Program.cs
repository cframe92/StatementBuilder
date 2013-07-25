using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;
using System.Xml.Serialization;
using iTextSharp.text;
using iTextSharp.text.pdf;
using XMLTest.Classes;
using System.Diagnostics;
using System.Reflection;



namespace XMLTest
{
    class Program
    {
        static void Main(string[] args)
        {
            
            Console.WriteLine("Copy the Statements extract to " + Configuration.GetSymitarStatementDataFilePath() + " then press any key to continue.");
            Console.ReadLine();
            Console.WriteLine("Copy the MoneyPerks extract to " + Configuration.GetMoneyPerksFilePath() + " then press any key to continue.");
            Console.ReadLine();
            Console.WriteLine("Ensure the following path exists: " + Configuration.GetStatementsOutputPath() + " then press any key to continue.");
            Console.ReadLine();
            Console.WriteLine("Ensure the following path exists: " + Configuration.GetErrorLogOutputPath() + " then press any key to continue.");
            Console.ReadLine();
            Stopwatch stopwatch = Stopwatch.StartNew();
            LogWriter = new StreamWriter(Configuration.GetErrorLogOutputPath() + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".log");
            CleanStatementsOutputPath();

            StatementBuilder.BuildMoneyPerksStatements();
           
            //BuildMemberStatements();
            LogWriter.Close();
            stopwatch.Stop();
            Console.WriteLine(StatementBuilder.GetNumberOfStatementsBuilt() + " statements produced in " + stopwatch.Elapsed.TotalMinutes.ToString("N") + " minutes.");
        
        }
        static void CleanStatementsOutputPath()
        {
            Console.WriteLine("I am going to delete all of the files in " + Configuration.GetStatementsOutputPath() + ".  Press any key to continue.");
            Console.ReadLine();
            Console.WriteLine("Deleting files from " + Configuration.GetStatementsOutputPath() + "...");
            System.IO.DirectoryInfo dirInfo = new DirectoryInfo(Configuration.GetStatementsOutputPath());

            foreach (FileInfo fileInfo in dirInfo.GetFiles())
            {
                fileInfo.Delete();
            }
        }

  
        static StreamWriter LogWriter
        {
            get;
            set;
        }

  
    }

    }


