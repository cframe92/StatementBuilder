using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Xml;
using System.Xml.Serialization;
using System.Configuration;
using System.Diagnostics;
using System.Reflection;

namespace XMLTest.Classes
{
    public class StatementBuilder
    {

        public static void BuildMoneyPerksStatements()
        {
            Console.WriteLine("Processing MoneyPerks file " + Configuration.GetMoneyPerksFilePath() + "...");
            StreamReader reader = new StreamReader(Configuration.GetMoneyPerksFilePath());
            MoneyPerksStatements = new Dictionary<string, MoneyPerksStatement>();

            while (!reader.EndOfStream)
            {
                List<string> fields = CSVParser(reader.ReadLine());

                if (fields.Count == MONEYPERKS_TRANSACTION_RECORD_FIELD_COUNT)
                {
                    MoneyPerksStatement moneyPerksStatement = null;
                    MoneyPerksTransaction transaction = new MoneyPerksTransaction();
                    string accountNumber = fields[0];
                    DateTime date = ParseDate(fields[1]);
                    string description = fields[2];
                    int amount = ParseMoneyPerksAmount(fields[3]);
                    int balance = ParseMoneyPerksAmount(fields[4]);

                    if (MoneyPerksStatements.ContainsKey(accountNumber))
                    {
                        moneyPerksStatement = MoneyPerksStatements[accountNumber];
                    }
                    else
                    {
                        moneyPerksStatement = new MoneyPerksStatement(accountNumber);
                        MoneyPerksStatements.Add(accountNumber, moneyPerksStatement);
                    }

                    if (description == "Beginning Balance")
                    {
                        moneyPerksStatement.BeginningBalance = balance;
                    }
                    else if (description == "Ending Balance")
                    {
                        moneyPerksStatement.EndingBalance = balance;
                    }
                    else
                    {
                        transaction.Date = date;
                        transaction.Description = description;
                        transaction.Amount = amount;
                        transaction.Balance = balance;
                        moneyPerksStatement.Transactions.Add(transaction);
                        MoneyPerksStatements[accountNumber] = moneyPerksStatement;
                    }
                }
            }

            reader.Close();
            reader.Dispose();
            Console.WriteLine("Done processing MoneyPerks file");

            StatementBuilder b = new StatementBuilder();
            b.DeserializeObject("C:\\Samples\\XML\\SampleStatementScrubbed.xml");
           
            
            
        }

        public void DeserializeObject(string filename)
        {
            AdvertisementTop = new Advertisement[MAX_RELATIONSHIP_BASED_LEVELS];

            for(int i = 0; i < AdvertisementTop.Count(); i++)
            {
                AdvertisementTop[i] = new Advertisement();
            }


            XmlSerializer serializer = new XmlSerializer(typeof(statementProduction));

            FileStream fs = new FileStream(filename, FileMode.Open);
            XmlReader reader = new XmlTextReader(fs);

            statementProduction s;

            s = (statementProduction)serializer.Deserialize(reader);

            StatementBuilder.Build(s, null);
            Console.ReadLine();
            reader.Close();
            Console.WriteLine("Done processing statement file.");
            Console.ReadLine();
        }

        public static void Build(statementProduction statement, MoneyPerksStatement moneyPerksStatement)
        {
            

                using (FileStream outputStream = File.Create("C:\\" + TEMP_FILE_NAME))
                {
                    CreateFirstPage(statement, outputStream);
                    AddMoneyPerksSummary(moneyPerksStatement);
                    AddBottomAdvertising(statement);
                    Doc.Close();
                }

            AddPageNumbersAndDisclosures(statement); // Re-opens document to overlay page numbers
/*
            if (File.Exists("c:\\" + TEMP_FILE_NAME))
            {
                File.Delete("c:\\" + TEMP_FILE_NAME);
            }
*/
            NumberOfStatementsBuilt++;
        }

        static void CreateFirstPage(statementProduction statement, FileStream outputStream)
        {
            DateTime statementStart = statement.envelope[0].statement.beginningStatementDate;
            DateTime statementEnd = statement.prologue.statementEndingDate;
            int shareCount = statement.epilogue.shareCount;
            int loanCounts = statement.epilogue.loanCount;
            int envelopeCount = statement.epilogue.envelopeCount;
            int accountCount = statement.epilogue.accountCount;

            int subAccountCount = 0;
            statementProductionEnvelopeStatementAccountSubAccount[] subaccounts;


            //Adds first page template to statement
            using (FileStream templateInputStream = File.Open(Configuration.GetStatementTemplateFirstPageFilePath(), FileMode.Open))
            {
                PdfReader reader = new PdfReader(templateInputStream);
                Doc = new Document(reader.GetPageSize(1));
                Writer = PdfWriter.GetInstance(Doc, outputStream);
                StatementPageEvent pageEvent = new StatementPageEvent();
                Writer.PageEvent = pageEvent;
                Writer.SetFullCompression();
                Doc.Open();
                PdfContentByte contentByte = Writer.DirectContent;
                PdfImportedPage page = Writer.GetImportedPage(reader, 1);
                Doc.NewPage();
                contentByte.AddTemplate(page, 0, 0);
            }

            AddStatementHeading("Statement  of  Accounts", 409, 0);
            AddStatementHeading(statementStart.ToString("MMM  dd,  yyyy") + "  thru  " + statementEnd.ToString("MMM  dd,  yyyy"), 385, 6f);


            //Set up Address and balances section
            PdfPTable addressAndBalancesTable = new PdfPTable(2);
            float[] addressAndBalancesTableWidths = new float[] { 50f, 50f };
            addressAndBalancesTable.SetWidthPercentage(addressAndBalancesTableWidths, Doc.PageSize);
            addressAndBalancesTable.TotalWidth = 612f;
            addressAndBalancesTable.LockedWidth = true;


            //Should we iterate and add a statement heading account number for every one or just the first account #?
            
                
                //addBasicAccountDetails(statement.envelope[j].statement.account);
            


            //Count the number of sub accounts in each envelope
            for (int i = 0; i < envelopeCount; i++)
            {
                if (statement.envelope[i].statement.account.accountNumber > 4)
                {
                    AddStatementHeading("Account  Number:        ******" + statement.envelope[i].statement.account.accountNumber.Equals("******".Length), 385, 6f);
                }

                if (statement.envelope[i].address != null)
                {
                    AddAddress(statement.envelope[i], ref addressAndBalancesTable);
                }

                subaccounts = statement.envelope[i].statement.account.subAccount;

                foreach (statementProductionEnvelopeStatementAccountSubAccount sub in subaccounts)
                {
                    subAccountCount++;
                }

                sortSubAccounts(loanCounts, accountCount, subAccountCount, subaccounts, ref addressAndBalancesTable);
                i++;

                if (i < envelopeCount)
                {
                    subAccountCount = 0;
                    subaccounts = statement.envelope[i].statement.account.subAccount;

                    foreach (statementProductionEnvelopeStatementAccountSubAccount sb in subaccounts)
                    {
                        subAccountCount++;
                    }
                }

                sortSubAccounts(loanCounts, accountCount, subAccountCount, subaccounts, ref addressAndBalancesTable);
            }

            Doc.Add(addressAndBalancesTable);
            //AddTopAdvertising(statement);
            AddInvisibleAccountNumber(statement);

            Console.ReadLine();


        }

        /*
        public static void addBasicAccountDetails(statementProductionEnvelopeStatementAccount account)
        {
            statementProductionEnvelopeStatementAccount tAccount;
            tAccount = account;

            if(tAccount.accountNumber > 0)
            {

            }

        }
        */
        static void sortSubAccounts(int loanCount, int accountCount, int count, statementProductionEnvelopeStatementAccountSubAccount[] subaccount, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountLoan loan;
            statementProductionEnvelopeStatementAccountSubAccountShare share;
            string subAccountCategory;


            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk("Account  Balances  at  a  Glance:", GetBoldFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 81;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            PdfPTable balancesTable = new PdfPTable(2);
            float[] tableWidths = new float[] { 60f, 40f };
            balancesTable.SetWidthPercentage(tableWidths, Doc.PageSize);
            balancesTable.TotalWidth = 300f;
            balancesTable.LockedWidth = true;
            

            for (int i = 0; i < count; i++)
            {
                if (subaccount[i].loan != null)
                {
                    loan = subaccount[i].loan;
                    AddYtdSummaries(loanCount, accountCount, subaccount[i]);
                    subAccountCategory = loan.category.Value;
                    sortLoanCategories(subaccount[i].loan, ref balancesTable);
                }
                else if (subaccount[i].share != null)
                {
                    share = subaccount[i].share;
                    AddYtdSummaries(loanCount, accountCount, subaccount[i]);
                    subAccountCategory = (string)share.category.Value;
                    sortShareCategories(subaccount[i].share, ref balancesTable);
                }
            }
        }

       
        

        public static void sortLoanCategories(statementProductionEnvelopeStatementAccountSubAccountLoan category, ref PdfPTable table)
        {
            string loanCategory;
            loanCategory = (string)category.category.Value;
           

            switch (loanCategory)
            {
                case "Closed end":
                    addHeaderBalances(null, category, "Loan", ref table);
                    AddClosedEndLoan(category, ref table);
                    break;
                case "Open end":
                    addHeaderBalances(null, category, "Loan", ref table);
                    AddOpenEndLoan(category, ref table);
                    break;
                case "Line of credit":
                    addHeaderBalances(null, category, "Loan", ref table);
                    AddLineOfCreditLoan(category, ref table);
                    break;
                case "Credit card":
                    addHeaderBalances(null, category, "Loan", ref table);
                    AddLineOfCreditLoan(category, ref table);
                    break;
                default:
                    Console.WriteLine("Invalid category");
                    break;
            }
        }

       

        public static void sortShareCategories(statementProductionEnvelopeStatementAccountSubAccountShare category, ref PdfPTable table)
        {
            string shareCategory;
            shareCategory = (string)category.category.Value;

            switch (shareCategory)
            {
                case "Share":
                    addHeaderBalances(category, null, "Total Savings: ", ref table);
                    AddSavingsAccounts(category, ref table);
                    break;
                case "Draft":
                    addHeaderBalances(category, null, "Total Checking: ", ref table);
                    AddCheckingAccounts(category, ref table);
                    break;
                case "Club":
                    addHeaderBalances(category, null, "Total Clubs: ", ref table);
                    AddClubAccounts(category, ref table);
                    break;
                case "Certificate":
                    addHeaderBalances(category, null, "Total Certificates: ", ref table);
                    AddCertificateAccounts(category, ref table);
                    break;
                default:
                    Console.WriteLine("Invalid Category");
                    break;
            }
        }

        static void AddClosedEndLoan(statementProductionEnvelopeStatementAccountSubAccountLoan closedEndLoan, ref PdfPTable table)
        {
            int transactionCount = 0;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransaction[] transactions;
            statementProductionEnvelopeStatementAccountSubAccountLoanInterestCharge interestCharge;
            //statementProductionEnvelopeStatementAccountSubAccountLoanBalanceComputationMethod method;
            int i = 1;
            decimal annualPercentageRate = closedEndLoan.beginning.annualRate;
            decimal dailyPeriodicRate = closedEndLoan.beginning.dailyPeriodicRate;

            AddSectionHeading("LOAN ACCOUNTS");
            AddAccountSubHeading(closedEndLoan.description, i > 0);
            
            
            if(closedEndLoan.maturityDateSpecified)
            {
                AddLoanTransactionsFooter("Closing Date of Billing Cycle " + closedEndLoan.endingStatementDate.ToString("MM/dd/yyyy") + "\n" +
                        "** INTEREST CHARGE CALCULATION: The balance used to compute interest charges is the unpaid balance each day after payments and credits to that balance have been subtracted and any additions to the balance have been made.");
                AddFeeSummary(closedEndLoan);
                AddInterestChargedSummary(closedEndLoan);
                AddLoanClosed(closedEndLoan, ref table);
            }
            AddYearToDateTotals(closedEndLoan);

            //Verify there are loan transactions
            if (closedEndLoan.transaction != null)
            {
                transactions = closedEndLoan.transaction;

                foreach (statementProductionEnvelopeStatementAccountSubAccountLoanTransaction t in transactions)
                {
                    transactionCount++;
                }

                sortLoanTransactions(transactions, ref table);
            }
            else
            {
                NoTransactionsThisPeriodMessage(ref table);
            }

            //Check if there are interest charges
            if (closedEndLoan.interestCharge != null)
            {
                interestCharge = closedEndLoan.interestCharge;
                addLoanInterest(interestCharge);
            }

        }

        static void AddOpenEndLoan(statementProductionEnvelopeStatementAccountSubAccountLoan openEndLoan, ref PdfPTable table)
        {
            int transactionCount = 0;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransaction[] transactions;
            statementProductionEnvelopeStatementAccountSubAccountLoanInterestCharge interestCharge;
            //statementProductionEnvelopeStatementAccountSubAccountLoanBalanceComputationMethod method;
            int i = 1;
            decimal annualPercentageRate = openEndLoan.beginning.annualRate;
            decimal dailyPeriodicRate = openEndLoan.beginning.dailyPeriodicRate;

            AddSectionHeading("LOAN ACCOUNTS");
            AddAccountSubHeading(openEndLoan.description, i > 0);

            if (openEndLoan.maturityDateSpecified)
            {
                AddLoanTransactionsFooter("Closing Date of Billing Cycle " + openEndLoan.endingStatementDate.ToString("MM/dd/yyyy") + "\n" +
                        "** INTEREST CHARGE CALCULATION: The balance used to compute interest charges is the unpaid balance each day after payments and credits to that balance have been subtracted and any additions to the balance have been made.");
                AddFeeSummary(openEndLoan);
                AddInterestChargedSummary(openEndLoan);
                AddLoanClosed(openEndLoan, ref table);
            }
            AddYearToDateTotals(openEndLoan);


            //Verify the loan has transactions associated with it
            if(openEndLoan.transaction != null)
            {
                transactions = openEndLoan.transaction;

                foreach(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction t in transactions)
                {
                    transactionCount++;
                }

                sortLoanTransactions(transactions, ref table);
            }
            else
            {
                NoTransactionsThisPeriodMessage(ref table);
            }


            //Verify the loan has interest charges
            if(openEndLoan.interestCharge != null)
            {
                interestCharge = openEndLoan.interestCharge;
                addLoanInterest(interestCharge);
            }
        }

        static void AddCreditCardLoan(statementProductionEnvelopeStatementAccountSubAccountLoan creditCardLoan, ref PdfPTable table)
        {
            int transactionCount = 0;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransaction[] transactions;
            statementProductionEnvelopeStatementAccountSubAccountLoanInterestCharge interestCharge;
            //statementProductionEnvelopeStatementAccountSubAccountLoanBalanceComputationMethod method;
            int i = 1;
            decimal annualPercentageRate = creditCardLoan.beginning.annualRate;
            decimal dailyPeriodicRate = creditCardLoan.beginning.dailyPeriodicRate;

            AddSectionHeading("LOAN ACCOUNTS");
            AddAccountSubHeading(creditCardLoan.description, i > 0);

            if (creditCardLoan.maturityDateSpecified)
            {
                AddLoanTransactionsFooter("Closing Date of Billing Cycle " + creditCardLoan.endingStatementDate.ToString("MM/dd/yyyy") + "\n" +
                        "** INTEREST CHARGE CALCULATION: The balance used to compute interest charges is the unpaid balance each day after payments and credits to that balance have been subtracted and any additions to the balance have been made.");
                AddFeeSummary(creditCardLoan);
                AddInterestChargedSummary(creditCardLoan);
                AddLoanClosed(creditCardLoan, ref table);
            }
            AddYearToDateTotals(creditCardLoan);

            //Verify the loan has transactions associated with it
            if(creditCardLoan.transaction != null)
            {
                transactions = creditCardLoan.transaction;

                foreach(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction t in transactions)
                {
                    transactionCount++;
                }

                sortLoanTransactions(transactions, ref table);
            }
            else
            {
                NoTransactionsThisPeriodMessage(ref table);
            }


            //Verify the loan has interest
            if(creditCardLoan.interestCharge != null)
            {
                interestCharge = creditCardLoan.interestCharge;
                addLoanInterest(interestCharge);
            }

        }

        static void AddLineOfCreditLoan(statementProductionEnvelopeStatementAccountSubAccountLoan lineOfCreditLoan, ref PdfPTable table)
        {
            int transactionCount = 0;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransaction[] transactions;
            statementProductionEnvelopeStatementAccountSubAccountLoanInterestCharge interestCharge;

            int i = 1;
            decimal annualPercentageRate = lineOfCreditLoan.beginning.annualRate;
            decimal dailyPeriodicRate = lineOfCreditLoan.beginning.dailyPeriodicRate;

            AddSectionHeading("LOAN ACCOUNTS");
            AddAccountSubHeading(lineOfCreditLoan.description, i > 0);

            if (lineOfCreditLoan.maturityDateSpecified)
            {
                AddLoanTransactionsFooter("Closing Date of Billing Cycle " + lineOfCreditLoan.endingStatementDate.ToString("MM/dd/yyyy") + "\n" +
                        "** INTEREST CHARGE CALCULATION: The balance used to compute interest charges is the unpaid balance each day after payments and credits to that balance have been subtracted and any additions to the balance have been made.");
                AddFeeSummary(lineOfCreditLoan);
                AddInterestChargedSummary(lineOfCreditLoan);
                AddLoanClosed(lineOfCreditLoan, ref table);
            }
            AddYearToDateTotals(lineOfCreditLoan);


            //Verify the loan has transactions associated with it
            if (lineOfCreditLoan.transaction != null)
            {
                transactions = lineOfCreditLoan.transaction;

                foreach (statementProductionEnvelopeStatementAccountSubAccountLoanTransaction t in transactions)
                {
                    transactionCount++;
                }

                sortLoanTransactions(transactions, ref table);
            }
            else
            {
                NoTransactionsThisPeriodMessage(ref table);
            }

            //Verify the loan has interest
            if (lineOfCreditLoan.interestCharge != null)
            {
                interestCharge = lineOfCreditLoan.interestCharge;
                addLoanInterest(interestCharge);
            }

            

        }

        static void AddCheckingAccounts(statementProductionEnvelopeStatementAccountSubAccountShare checking, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountShare tChecking;
            statementProductionEnvelopeStatementAccountSubAccountShareTransaction[] checkingTransactions;
            statementProductionEnvelopeStatementAccountSubAccountSharePerson additionalName;
            tChecking = checking;
            int i = 1;
            int transactionCount = 0;


            AddSectionHeading("CHECKING ACCOUNTS");
            AddAccountSubHeading(tChecking.description, i > 0);

            //Verify the checking account has transactions associated with it
            if (tChecking.transaction != null)
            {
                checkingTransactions = tChecking.transaction;

                foreach (statementProductionEnvelopeStatementAccountSubAccountShareTransaction t in checkingTransactions)
                {
                    transactionCount++;
                }

                sortShareTransactions(checkingTransactions, ref table);
            }
            else
            {
                NoTransactionsThisPeriodMessage(ref table);
            }

            if (tChecking.closeDateSpecified == true)
            {
                AddShareClosed(tChecking, ref table);
            }

            if(tChecking.person != null)
            {
                additionalName = tChecking.person;
                sortAdditionalNames(additionalName);
            }

            //if the total overdraftfree amount ytd + totalreturneditemfee amount ytd is greater than 0
           // if ((account.TotalOverdraftFee.AmountYtd + account.TotalReturnedItemFee.AmountYtd) > 0)
           // {
           //     AddTotalFees(account);
           // }



            // Adds APR
            /*
                    if (statement.envelope[i].statement.account.subAccount[i].loan.beginning.annualRate > 0)
                    {
                        PdfPTable table = new PdfPTable(1);
                        table.TotalWidth = 525f;
                        table.LockedWidth = true;
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("Annual Percentage Yield Earned " + statement.envelope[i].statement.account.subAccount[i].loan.beginning.annualRate.ToString("N3") + "% from " + account.AnnualPercentageRate.BeginningDate.ToString("MM/dd/yyyy") + " through " + account.AnnualPercentageRate.EndingDate.ToString("MM/dd/yyyy"), GetBoldItalicFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        //p.IndentationLeft = 70;
                        cell.AddElement(p);
                        cell.PaddingTop = -4f;
                        cell.BorderWidth = 0f;
                        table.AddCell(cell);
                        Doc.Add(table);
                    }
             */
   
        }

        static void AddSavingsAccounts(statementProductionEnvelopeStatementAccountSubAccountShare savings, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountShare tSavings;
            statementProductionEnvelopeStatementAccountSubAccountSharePerson additionalName;
            statementProductionEnvelopeStatementAccountSubAccountShareTransaction[] savingsTransaction;
            tSavings = savings;
            int i = 1;
            int transactionCount = 0;
            
            AddSectionHeading("SAVINGS ACCOUNTS");
            AddAccountSubHeading(tSavings.description, i > 0);

            //Verify the savings account has transactions
            if (tSavings.transaction != null)
            {
                savingsTransaction = tSavings.transaction;

                foreach (statementProductionEnvelopeStatementAccountSubAccountShareTransaction t in savingsTransaction)
                {
                    transactionCount++;

                }
                sortShareTransactions(savingsTransaction, ref table);
            }
            else
            {
                NoTransactionsThisPeriodMessage(ref table);
            }

            if (tSavings.closeDateSpecified == true)
            {
                AddShareClosed(tSavings, ref table);
            }
            if (tSavings.person != null)
            {
                additionalName = tSavings.person;
                sortAdditionalNames(additionalName);
            }

        }


        static void AddClubAccounts(statementProductionEnvelopeStatementAccountSubAccountShare club, ref PdfPTable table)
        {

            statementProductionEnvelopeStatementAccountSubAccountShare tClub;
            statementProductionEnvelopeStatementAccountSubAccountShareTransaction[] clubTransactions;
            statementProductionEnvelopeStatementAccountSubAccountSharePerson additionalName;
            tClub = club;
            int i = 1;
            int transactionCount = 0;
            

            AddSectionHeading("CLUB ACCOUNTS");
            AddAccountSubHeading(tClub.description, i > 0);

            //Verify the club account has transactions
            if (tClub.transaction != null)
            {
                clubTransactions = tClub.transaction;

                foreach (statementProductionEnvelopeStatementAccountSubAccountShareTransaction t in clubTransactions)
                {
                    transactionCount++;
                }

                sortShareTransactions(clubTransactions, ref table);
            }
            else
            {
                NoTransactionsThisPeriodMessage(ref table);
            }

            /*
            if (tClub.dividendYTDSpecified == true)
            {
                Console.WriteLine("Dividend YTD: {0}", tClub.dividendYTD);
            }
            */
            if (tClub.closeDateSpecified == true)
            {
                AddShareClosed(tClub, ref table);
            }

            if(tClub.person != null)
            {
                additionalName = tClub.person;
                sortAdditionalNames(additionalName);
            }

        }

        static void AddCertificateAccounts(statementProductionEnvelopeStatementAccountSubAccountShare certificate, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountShare tCertificate;
            statementProductionEnvelopeStatementAccountSubAccountShareTransaction[] certificateTransactions;
            statementProductionEnvelopeStatementAccountSubAccountSharePerson additionalName;
            tCertificate = certificate;
            int i = 1;
            int transactionCount = 0;

            AddSectionHeading("CERTIFICATE ACCOUNTS");
            string descriptionAndMaturityDate = tCertificate.description + "   Maturity Date - " + tCertificate.maturityDate.ToString("MMM dd, yyyy");
            AddAccountSubHeading(descriptionAndMaturityDate, i > 0);

            //verify the certificate has transactions
            if (tCertificate.transaction != null)
            {
                certificateTransactions = tCertificate.transaction;

                foreach (statementProductionEnvelopeStatementAccountSubAccountShareTransaction t in certificateTransactions)
                {
                    transactionCount++;
                }
                sortShareTransactions(certificateTransactions, ref table);
            }
            else
            {
                NoTransactionsThisPeriodMessage(ref table);
            }

            if (tCertificate.closeDateSpecified == true)
            {
                AddShareClosed(tCertificate, ref table);
            }
            if (tCertificate.person != null)
            {
                additionalName = tCertificate.person;
                sortAdditionalNames(additionalName);
            }

        }

        

        public static void sortLoanTransactions(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction[] loantransaction, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory loanCategory;

            for (int i = 0; i < loantransaction.Length;i++ )
            {
                decimal itemLength = loantransaction[i].Items.Length;

                for(int j = 0; j < itemLength; j++)
                {
                    string item = loantransaction[i].ItemsElementName[j].ToString();

                    switch(item)
                    {
                        case "category":
                            loanCategory = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory)loantransaction[i].Items[j];
                            sortLoanTransactionCategory(loanCategory.Value, loantransaction[i], ref table);
                            break;
                        default:
                            break;
                    }
                }
            }
        }


        public static void sortShareTransactions(statementProductionEnvelopeStatementAccountSubAccountShareTransaction[] transactionlist, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionCategory shareCategory;


            for (int i = 0; i < transactionlist.Length;i++ )
            {
                decimal itemLength = transactionlist[i].Items.Length;

                for(int j = 0; j < itemLength; j++)
                {
                    string item = transactionlist[i].ItemsElementName[j].ToString();

                    switch(item)
                    {
                        case "category":
                            shareCategory = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionCategory)transactionlist[i].Items[j];
                            sortShareTransactionCategory(shareCategory.Value, transactionlist[i], ref table);
                            break;
                        default:
                            break;
                    }
                }
            }

        }

        public static void sortShareTransactionCategory(string value, statementProductionEnvelopeStatementAccountSubAccountShareTransaction transaction, ref PdfPTable table)
        {
            switch(value)
            {
                case "Deposit":
                    sortDepositSource(transaction, ref table);
                    break;
                case "Withdrawal":
                    sortWithdrawalSource(transaction, ref table);
                    break;
                case "Comment":
                    addShareCommentTransaction(transaction, ref table);
                    break;
                default:
                    break;
            }
        }

        public static void sortLoanTransactionCategory(string value, statementProductionEnvelopeStatementAccountSubAccountLoanTransaction t, ref PdfPTable table)
        {

            switch (value)
            {
                case "Payment":
                    addPayment(t, ref table);
                    break;
                case "Advance":
                    AddAdvances(t, ref table);
                    break;
                case "Refinance":
                    addRefinance(t, ref table);
                    break;
                case "New loan":
                    addNewLoan(t, ref table);
                    break;
                case "Comment":
                    addLoanCommentTransaction(t, ref table);
                    break;
                default:
                    break;
            }
        }


        public static void sortDepositSource(statementProductionEnvelopeStatementAccountSubAccountShareTransaction deposit, ref PdfPTable table)
        {
            decimal transactionItemLength = deposit.Items.Length;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource depositSource;

            for(int i = 0; i < transactionItemLength; i++)
            {
                string item = deposit.ItemsElementName[i].ToString();

                switch(item)
                {
                    case "source":
                        depositSource = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource)deposit.Items[i];
                        if (depositSource.Value == "ATM")
                            AddAtmDeposit(deposit, ref table);
                        else
                            AddDeposit(deposit);
                        break;
                    default:
                        break;
                }
            }
        }

        public static void sortWithdrawalSource(statementProductionEnvelopeStatementAccountSubAccountShareTransaction withdrawal, ref PdfPTable table)
        {
            decimal transactionItemLength = withdrawal.Items.Length;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource withdrawalSource;

            for (int i = 0; i < transactionItemLength; i++)
            {
                string item = withdrawal.ItemsElementName[i].ToString();

                switch (item)
                {
                    case "source":
                        withdrawalSource = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource)withdrawal.Items[i];
                        if (withdrawalSource.Value == "ATM")
                            AddAtmWithdrawal(withdrawal, ref table);
                        else
                            AddWithdrawal(withdrawal);
                        break;
                    default:
                        break;
                }
            }
        }

        public static void AddAtmDeposit(statementProductionEnvelopeStatementAccountSubAccountShareTransaction atmDeposit, ref PdfPTable table)
        {
            int counter = 0;
            int atmDepositItems = atmDeposit.Items.Length;
            int result;

            while(counter < atmDepositItems)
            {
                string atmDepositItem = atmDeposit.ItemsElementName[counter].ToString();
                result = addShareTransactionElements(counter, atmDeposit.ItemsElementName[counter]);

                if(atmDeposit.Items[result] != null)
                {
                    if(result == counter)
                    {
                        switch(atmDepositItem)
                        {
                            case "postingDate":
                                break;
                            default:
                                Console.WriteLine("ATM DEPOSIT ITEM: {0}", atmDeposit.ItemsElementName[counter]);
                                break;
                        }
                    }
                    else
                    {
                        Console.WriteLine("ATM DEPOSIT COUNTER ITEM: {0}", atmDeposit.ItemsElementName[counter]);
                        Console.WriteLine("ATM DEPOSIT RESULT ITEM :{0}", atmDeposit.ItemsElementName[result]);
                    }
                }
                counter++;
            }
        }

        public static void AddAtmWithdrawal(statementProductionEnvelopeStatementAccountSubAccountShareTransaction atmWithdrawal, ref PdfPTable table)
        {
            int counter = 0;
            int atmWithdrawalItems = atmWithdrawal.Items.Length;
            int result;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionCategory category;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource source;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionTransferOption option;
            //statementProductionEnvelopeStatementAccountSubAccountShareTransactionSubCategory subcategory;

            while(counter < atmWithdrawalItems)
            {
                string atmWithdrawalItem = atmWithdrawal.ItemsElementName[counter].ToString();
                result = addShareTransactionElements(counter, atmWithdrawal.ItemsElementName[counter]);

                if(atmWithdrawal.Items[result] != null)
                {
                    if(result == counter)
                    {
                        switch(atmWithdrawalItem)
                        {
                            case "postingDate":
                                break;
                            case "grossAmount":
                                break;
                            case "principal":
                                break;
                            case "newBalance":
                                break;
                            case "terminalLocation":
                                break;
                            case "monetarySerial":
                                break;
                            case "transactionSerial":
                                break;
                            case "accountNumber":
                                break;
                            case "terminalId":
                                break;
                            case "terminalCity":
                                break;
                            case "terminalState":
                                break;
                            case "merchantName":
                                break;
                            case "merchantType":
                                break;
                            case "transactionReference":
                                break;
                            case "transactionDate":
                                break;
                            case "maskedCardNumber":
                                break;
                            default:
                                Console.WriteLine("ATM WITHDRAWAL ITEM: {0}", atmWithdrawal.ItemsElementName[counter]);
                                break;
                        }
                    }
                    else
                    {
                        switch(atmWithdrawalItem)
                        {
                            case "category":
                                category = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionCategory)atmWithdrawal.Items[counter];
                                break;
                            case "source":
                                source = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource)atmWithdrawal.Items[counter];
                                break;
                            case "transferOption":
                                option = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionTransferOption)atmWithdrawal.Items[counter];
                                Console.WriteLine("ATM WITH OPT: {0}", option.Value);
                                break;
                        }
                    }
                }
                counter++;
            }

        }
        /*
        public static void AddLoanTransaction(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction transaction, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory loanTransactionCategory;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource loanTransactionSource;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionTransferOption loanTransferOption;
            int loanTransactionItems;
            loanTransactionItems = transaction.Items.Length;
            int counter = 0;
            int returnedValue;
            string description = string.Empty;
            PdfPCell cell = new PdfPCell();
            decimal newBal = 0;
            
            
            //addEndingLoanBalance(transaction, ref table);
            

            while(counter < loanTransactionItems)
            {
                string loanItem = transaction.ItemsElementName[counter].ToString();

                returnedValue = addLoanTransactionElements(counter, transaction.ItemsElementName[counter]);

                if((transaction.Items[returnedValue] != null)&& (returnedValue == counter))
                {
                    switch(loanItem)
                    {
                        case "postingDate":
                            Chunk chunk = new Chunk(transaction.Items[returnedValue].ToString(), GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            table.AddCell(cell);
                            Console.WriteLine("Loan Trans Posting Date: {0}", transaction.Items[returnedValue]);
                            break;
                        case "description":
                            description = (string)transaction.Items[returnedValue];

                            PdfPCell Pcell = new PdfPCell();
                            Chunk Dchunk = new Chunk(description, GetNormalFont(9f));
                            Dchunk.SetCharacterSpacing(0f);
                            Paragraph par = new Paragraph(Dchunk);
                            Pcell.AddElement(par);

                            Pcell.PaddingTop = -6f;
                            Pcell.BorderWidth = 0f;
                            table.AddCell(Pcell);
                            break;
                        case "newBalance":
                            newBal = (decimal)transaction.Items[returnedValue];
                            AddLoanAccountTransactionAmount(newBal, ref table); // Balance subject to interest rate**        
                            break;
                        case "monetarySerial":
                            break;
                        case "transactionSerial":
                            break;
                        case "grossAmount":
                            AddLoanAccountTransactionAmount((decimal)transaction.Items[returnedValue], ref table); // Amount
                            break;
                        case "principal":
                            AddLoanAccountTransactionAmount((decimal)transaction.Items[returnedValue], ref table); // Principal
                            break;
                        case "interest":
                            AddLoanAccountTransactionAmount((decimal)transaction.Items[returnedValue], ref table); // Interest Charged
                            break;
                        case "lateFee":
                            AddLoanAccountTransactionAmount((decimal)transaction.Items[returnedValue], ref table); // Late Fees
                            break;
                        case "transferId":
                            break;
                        case "transferIdCategory":
                            break;
                        default:
                            Console.WriteLine("LOAN ITEM: {0}", transaction.ItemsElementName[returnedValue]);
                            break;
                    }
                }

                else if(returnedValue != counter)
                {
                    if (returnedValue == 1)
                    {
                        string item = transaction.ItemsElementName[counter].ToString();

                        switch(item)
                        {
                            case "category":
                                loanTransactionCategory = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory)transaction.Items[counter];
                                //Console.WriteLine("Loan Trans Category: {0}", loanTransactionCategory.Value);
                                //sortLoanTransactionCategory(transaction, loanTransactionCategory, ref table);
                                break;
                            case "source":
                                loanTransactionSource = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource)transaction.Items[counter];
                                //Console.WriteLine("Loan Trans Source: {0}", loanTransactionSource.Value);
                                break;
                            case "transferOption":     
                                loanTransferOption = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionTransferOption)transaction.Items[counter];
                                //Console.WriteLine("Loan Transfer Option: {0}", loanTransferOption.Value);
                                break;
                            default:
                                //Console.WriteLine("LOAN ITEM: {0}", transaction.ItemsElementName[counter]);
                                break;
                            
                        }  
                    } 
                }
                counter++;
            }

            //if (transaction.Amount >= 0)
            //{
            //    AddAccountTransactionAmount(transaction.Amount, ref table); // Additions
            //    AddAccountTransactionAmount(0, ref table); // Subtractions
            //}
            //else
            //{
            //    AddAccountTransactionAmount(0, ref table); // Additions
            //    AddAccountTransactionAmount(transaction.Amount, ref table); // Subtractions
            //}

            AddAccountBalance(newBal, ref table);
            Doc.Add(table);

        }
        */
        /*
        public static void addShareTransaction(statementProductionEnvelopeStatementAccountSubAccountShareTransaction transactionlist, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionCategory shareTransactionCategory;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource shareTransactionSource;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionTransferOption shareTransferOption;
            int shareTransactionItems;
            int tracker = 0;
            shareTransactionItems = transactionlist.Items.Length;
            int returnedValue;
            string description = string.Empty;
            PdfPCell cell = new PdfPCell();
            decimal newBal = 0;
           

            while (tracker < shareTransactionItems)
            {

                returnedValue = addShareTransactionElements(tracker, (transactionlist.ItemsElementName[tracker]));

                if ((transactionlist.Items[returnedValue] != null) && (returnedValue == tracker))
                {

                    if(transactionlist.ItemsElementName[returnedValue].Equals("postingDate"))
                    {
                        Chunk chunk = new Chunk((transactionlist.Items[returnedValue].ToString()), GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        cell.AddElement(p);
                        cell.PaddingTop = -4f;
                        cell.BorderWidth = 0f;
                        table.AddCell(cell);
                    }

                    if(transactionlist.ItemsElementName[returnedValue].Equals("description"))
                    {
                        description = transactionlist.Items[returnedValue].ToString();
                        PdfPCell Dcell = new PdfPCell();
                        Chunk chunk = new Chunk(description, GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        Dcell.AddElement(p);


                        Dcell.PaddingTop = -4f;
                        Dcell.BorderWidth = 0f;
                        table.AddCell(Dcell);
                    }

                    if(transactionlist.ItemsElementName[returnedValue].Equals("newBalance"))
                    {
                        newBal = (decimal)transactionlist.Items[returnedValue];
                    }

                    if(transactionlist.ItemsElementName[returnedValue].Equals("grossAmount"))
                    {
                        decimal amount = (decimal)transactionlist.Items[returnedValue];

                        if(amount > 0)
                        {
                            AddAccountTransactionAmount(amount, ref table); // Additions
                            AddAccountTransactionAmount(0, ref table); // Subtractions
                        }
                        else
                        {
                            AddAccountTransactionAmount(0, ref table); // Additions
                            AddAccountTransactionAmount(amount, ref table); // Subtractions
                        }

                        AddAccountBalance(newBal, ref table);
                    }
                    
                }
                if (returnedValue != tracker)
                {
                    

                    if (returnedValue == 1)
                    {
                        string item = transactionlist.ItemsElementName[tracker].ToString();

                        if(item == "category")
                        {
                            shareTransactionCategory = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionCategory)transactionlist.Items[tracker];
                            //sortShareTransactionCategory(transactionlist, shareTransactionCategory, ref table);
                        }
                        if(item == "source")
                        {
                            shareTransactionSource = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource)transactionlist.Items[tracker];
                            Console.WriteLine("Share Trans Source: {0}", shareTransactionSource.Value);
                        }
                        if(item == "transferOption")
                        {
                            shareTransferOption = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionTransferOption)transactionlist.Items[tracker];
                            Console.WriteLine("Share Trans Option: {0}", shareTransferOption.Value);
                        }
      
                    }
                    else
                    {
                        Console.WriteLine("tracker value: {0}", tracker);
                       Console.WriteLine("tracker item:{0}", transactionlist.ItemsElementName[tracker]);
                        Console.WriteLine("Returned value: {0}", returnedValue);
                        Console.WriteLine("Returned item: {0}", transactionlist.Items[returnedValue]);
                        Console.WriteLine(transactionlist.Items[tracker]);
                    }
                }
                tracker++;
            }

            
            if (share.closeDateSpecified == false)
            {
                AddEndingBalance(share, ref table);
            }
            else
            {
                AddShareClosed(share, ref table);
            }
            
            Doc.Add(table);
        
        }
         */ 
        static void AddDeposit(statementProductionEnvelopeStatementAccountSubAccountShareTransaction deposit)
        {
            statementProductionEnvelopeStatementAccountSubAccountShareTransaction memberDeposit = new statementProductionEnvelopeStatementAccountSubAccountShareTransaction();

            int depositlength = deposit.Items.Length;
            int counter = 0;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionCategory depositcategory;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource depositsource;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionTransferOption depositOption;
            string item;
            DateTime depositDate = DateTime.Now;
            string depositDescription = string.Empty;
            decimal depositAmount = 0;

            int rowBreakPointIndex = (int)Math.Ceiling((double)NumberOfDeposits / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("DEPOSITS AND OTHER CREDITS");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            while (counter < depositlength)
            {
                item = deposit.ItemsElementName[counter].ToString();

                switch (item)
                {
                    case "category":
                        depositcategory = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionCategory)deposit.Items[counter];
                        break;
                    case "source":
                        depositsource = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource)deposit.Items[counter];
                        break;
                    case "postingDate":
                        depositDate = (DateTime)deposit.Items[counter];
                        break;
                    case "grossAmount":
                        depositAmount = (decimal)deposit.Items[counter];
                        break;
                    case "description":
                        depositDescription = (string)deposit.Items[counter];
                        break;
                    case "monetarySerial":
                        break;
                    case "transactionSerial":
                        break;
                    case "principal":
                        break;
                    case "newBalance":
                        break;
                    case "transferOption":
                        depositOption = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionTransferOption)deposit.Items[counter];
                        break;
                    case "transferId":
                        break;
                    case "transferIdCategory":
                        break;
                    case "apyeRate":
          
                        break;
                    case "apyeAverageBalance":
                
                        break;
                    case "apyePeriodStartDate":
                    
                        break;
                    case "apyePeriodEndDate":
                  
                        break;
                    case "accountNumber":
                        break;
                    case "terminalLocation":
                        break;
                    case "terminalId":
                        break;
                    case "terminalCity":
                        break;
                    case "transactionDate":
                        break;
                    case "transactionReference":
                        break;
                    case "terminalState":
                        break;
                    case "merchantName":
                        break;
                    case "merchantType":
                        break;
                    case "routingNumber":
                        break;
                    case "draftNumber":
                        break;
                    case "draftTracer":
                        break;
                    case "achCompanyName":
                        break;
                    case "achCompanyId":
                        break;
                    case "achCompanyEntryDescription":
                        break;
                    case "achCompanyDescriptiveDate":
                        break;
                    case "achOriginatingDFIId":
                        break;
                    case "achStandardEntryClassCode":
                        break;
                    case "achTransactionCode":
                        break;
                    case "achName":
                        break;
                    case "achIdentificationNumber":
                        break;
                    case "achTraceNumber":
                        break;
                    case "settlementDate":
                        break;
                    default:
                        Console.WriteLine("DEPOSIT ITEM:{0}", item);
                        break;
                }

                /*
                for (int i = 0; i < memberDeposit.Deposits.Count; i++ )
                {
                    if((i + 1) <= rowBreakPointIndex)
                    {
                        rows.Add(new SortTableRow());
                        rows[i].Column.Add(depositDate.ToString("MMM dd"));
                        rows[i].Column.Add(FormatAmount(depositAmount));
                        rows[i].Column.Add(depositDescription);
                        rows[i].Column.Add(string.Empty);
                        rows[i].Column.Add(string.Empty);
                        rows[i].Column.Add(string.Empty);
                    }
                    else
                    {
                        rows[i - rowBreakPointIndex].Column[3] = depositDate.ToString("MMM dd");
                        rows[i - rowBreakPointIndex].Column[4] = FormatAmount(memberDeposit.Deposits[i].Amount);
                        rows[i - rowBreakPointIndex].Column[5] = depositDescription;
                    }
                }
                */
                for (int i = 0; i < rows.Count; i++)
                {
                    AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                    AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                    AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                    AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                    AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                    AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                }

                Doc.Add(table);
                //memberDeposit.Deposits.Add(new Deposit(depositDescription, depositAmount, depositDate));
                /*
                if (memberDeposit.Deposits.Count() > 1)
                {
                    AddSortTableSubtotal(memberDeposit.Deposits.Count().ToString() + " Deposits and Other Credits for " + FormatAmount(memberDeposit.DepositsTotal));
                }
                 */


                counter++;
            }
        }




        public static int addShareTransactionElements(int position, ItemsChoiceType1 item)
        {
            switch (item)
            {
                case ItemsChoiceType1.transactionSerial:
                    return position;
                case ItemsChoiceType1.monetarySerial:
                    return position;
                case ItemsChoiceType1.postingDate:
                    return position;
                case ItemsChoiceType1.category:
                    return 1;
                case ItemsChoiceType1.source:
                    return 1;
                case ItemsChoiceType1.grossAmount:
                    return position;
                case ItemsChoiceType1.principal:
                    return position;
                case ItemsChoiceType1.newBalance:
                    return position;
                case ItemsChoiceType1.description:
                    return position;
                case ItemsChoiceType1.apyeRate:
                    return position;
                case ItemsChoiceType1.apyeAverageBalance:
                    return position;
                case ItemsChoiceType1.apyePeriodStartDate:
                    return position;
                case ItemsChoiceType1.apyePeriodEndDate:
                    return position;
                case ItemsChoiceType1.transactionAmount:
                    return position;
                case ItemsChoiceType1.availableAmount:
                    return position;
                case ItemsChoiceType1.certificatePenalty:
                    return position;
                case ItemsChoiceType1.achCompanyDescriptiveDate:
                    return position;
                case ItemsChoiceType1.achCompanyEntryDescription:
                    return position;
                case ItemsChoiceType1.achCompanyId:
                    return position;
                case ItemsChoiceType1.achCompanyName:
                    return position;
                case ItemsChoiceType1.achIdentificationNumber:
                    return position;
                case ItemsChoiceType1.achName:
                    return position;
                case ItemsChoiceType1.achOriginatingDFIId:
                    return position;
                case ItemsChoiceType1.achStandardEntryClassCode:
                    return position;
                case ItemsChoiceType1.achTraceNumber:
                    return position;
                case ItemsChoiceType1.achTransactionCode:
                    return position;
                case ItemsChoiceType1.adjustmentOption:
                    return position;
                case ItemsChoiceType1.draftNumber:
                    return position;
                case ItemsChoiceType1.draftTracer:
                    return position;
                case ItemsChoiceType1.settlementDate:
                    return position;
                case ItemsChoiceType1.routingNumber:
                    return position;
                case ItemsChoiceType1.accountNumber:
                    return position;
                case ItemsChoiceType1.feeClassification:
                    return position;
                case ItemsChoiceType1.maskedCardNumber:
                    return position;
                case ItemsChoiceType1.merchantName:
                    return position;
                case ItemsChoiceType1.merchantType:
                    return position;
                case ItemsChoiceType1.subCategory:
                    return position;
                case ItemsChoiceType1.terminalCity:
                    return position;
                case ItemsChoiceType1.terminalId:
                    return position;
                case ItemsChoiceType1.terminalLocation:
                    return position;
                case ItemsChoiceType1.terminalState:
                    return position;
                case ItemsChoiceType1.transactionDate:
                    return position;
                case ItemsChoiceType1.transactionReference:
                    return position;
                case ItemsChoiceType1.transferAccountNumber:
                    return position;
                case ItemsChoiceType1.transferId:
                    return position;
                case ItemsChoiceType1.transferIdCategory:
                    return position;
                case ItemsChoiceType1.transferIdDescription:
                    return position;
                case ItemsChoiceType1.transferName:
                    return position;
                case ItemsChoiceType1.transferOption:
                    return 1;
                default:
                    return 5;
            }
        }


        public static int addLoanTransactionElements(int position, ItemsChoiceType item)
        {
            switch (item)
            {
                case ItemsChoiceType.transactionSerial:
                    return position;
                case ItemsChoiceType.monetarySerial:
                    return position;
                case ItemsChoiceType.postingDate:
                    return position;
                case ItemsChoiceType.category:
                    return 1;
                case ItemsChoiceType.transferOption:
                    return 1;
                case ItemsChoiceType.source:
                    return 1;
                case ItemsChoiceType.grossAmount:
                    return position;
                case ItemsChoiceType.principal:
                    return position;
                case ItemsChoiceType.interest:
                    return position;
                case ItemsChoiceType.lateFee:
                    return position;
                case ItemsChoiceType.description:
                    return position;
                case ItemsChoiceType.newBalance:
                    return position;
                case ItemsChoiceType.transferIdCategory:
                    return position;
                case ItemsChoiceType.transferId:
                    return position;
                default:
                    return 5;
            }
        }
        
        static void addNewLoan(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction newLoan, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource newLoanSource;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory newLoanCategory;

            int newLoanLength = newLoan.Items.Length;
            int counter = 0;
            string item;

            while (counter < newLoanLength)
            {
                item = newLoan.ItemsElementName[counter].ToString();

                switch(item)
                {
                    case "source":
                        newLoanSource = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource)newLoan.Items[counter];
                        sortLoanSource(newLoan, newLoanSource, ref table);
                        break;
                    case "category":
                        newLoanCategory = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory)newLoan.Items[counter];
                        break;
                    case "transactionSerial":
                        break;
                    case "monetarySerial":
                        break;
                    case "description":
                        break;
                    case "grossAmount":
                        break;
                    default:
                        Console.WriteLine("NEW LOAN ITEM: {0}", newLoan.ItemsElementName[counter]);
                        break;
                }
               
                counter++;
            }
        }

        static void addPayment(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction payment, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource paymentSource;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory paymentCategory;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionTransferOption paymentTransferOption;

            int paymentLength = payment.Items.Length;
            int counter = 0;
            string item;

            NumberOfPayments++;

            int rowBreakPointIndex = (int)Math.Ceiling((double)NumberOfPayments / 2);
            PdfPTable Ptable = new PdfPTable(6);
            Ptable.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            Ptable.TotalWidth = 525f;
            Ptable.SetWidths(tableWidths);
            Ptable.LockedWidth = true;
            Ptable.SpacingBefore = 10f;


            AddSortTableHeading("LOAN PAYMENTS AND OTHER CREDITS");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref Ptable);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref Ptable);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref Ptable);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref Ptable);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref Ptable);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref Ptable);

   
            while (counter < paymentLength)
            {
                item = payment.ItemsElementName[counter].ToString();
                switch(item)
                {
                    case "source":
                        paymentSource = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource)payment.Items[counter];
                        sortLoanSource(payment, paymentSource, ref table);
                        break;
                    case "category":
                        paymentCategory = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory)payment.Items[counter];
                        break;
                    case "postingDate":
                 
                        break;
                    case "transactionSerial":
                
                        break;
                    case "monetarySerial":
                 
                        break;
                    case "grossAmount":
                        decimal amount = (decimal)payment.Items[counter];
                        break;
                    case "principal":
                       
                        break;
                    case "interest":
                     
                        break;
                    case "lateFee":
                       
                        break;
                    case "newBalance":
                    
                        break;
                    case "transferId":
                     
                        break;
                    case "transferIdCategory":
        
                        break;
                    case "transferOption":
                        paymentTransferOption = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionTransferOption)payment.Items[counter];
                        break;
                    default:
                        Console.WriteLine("PAYMENT ITEM :{0}", payment.ItemsElementName[counter]);
                        break;
                }
                        counter++;
          }
                 
        }

        static void sortLoanSource(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction transaction, statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource source, ref PdfPTable table)
        {
            string loanSource = source.Value;

            switch(loanSource)
            {
                case "Check":
                    AddCheck(transaction);
                    break;
                case "Cash and check":
                    break;
                case "Draft":
                    break;
                case "Credit or debit card":
                    break;
                case "Payroll":
                    break;
                case "Bill payment":
                    break;
                case "Audio":
                    break;
                case "Home banking":
                    break;
                case "Shared branch":
                    break;
                case "ATM":
                    break;
                case "ACH":
                    break;
                case "ACH origination":
                    break;
                case "POS":
                    break;
                case "Insurance":
                    break;
                case "Fee":
                    break;
                default:
                    break;

            }

        }

        static void AddCheck(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction transaction)
        {
            //Check newCheck = new Check();
            decimal checkAmount = 0;
            DateTime checkDate = DateTime.Now;
            int checkItemLength = transaction.Items.Length;
            string checkNumber = string.Empty;

            


            int rowBreakPointIndex = (int)Math.Ceiling((double)NumberOfChecks / 2);
            PdfPTable table = new PdfPTable(7);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 91, 43, 76, 105, 91, 43, 76 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;

            AddSortTableHeading("CHECK SUMMARY");

            AddSortTableTitle("Check #", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Date", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle(string.Empty, Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Check #", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Date", Element.ALIGN_RIGHT, ref table);


            for(int i = 0; i < checkItemLength; i++)
            {
                string checkItem = transaction.ItemsElementName[i].ToString();

                switch(checkItem)
                {
                    case "postingDate":
                        checkDate = (DateTime)transaction.Items[i];
                        break;
                    case "grossAmount":
                        checkAmount = (decimal)transaction.Items[i];
                        break;
                    default:
                        //Console.WriteLine("CHECKITEM :{0}", checkItem);
                        break;
                }
            }


            Check newCheck = new Check(checkNumber, checkAmount, checkDate);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < NumberOfChecks; i++)
            {
                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    //rows[i].Column.Add(statement.Checks[i].CheckNumber);
                    rows[i].Column.Add(FormatAmount(checkAmount));
                    rows[i].Column.Add(checkDate.ToString("MMM dd"));
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    //rows[i - rowBreakPointIndex].Column[4] = statement.Checks[i].CheckNumber;
                    rows[i - rowBreakPointIndex].Column[5] = FormatAmount(checkAmount);
                    rows[i - rowBreakPointIndex].Column[6] = checkDate.ToString("MMM dd");
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Check #
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_RIGHT, ref table); // Adds Date
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_RIGHT, ref table); // Empty column title
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_LEFT, ref table);  // Adds Check #
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[6], Element.ALIGN_RIGHT, ref table); // Adds Date
            }

            Doc.Add(table);
            AddChecksFootnote();

            if (NumberOfChecks > 1)
            {
               // AddSortTableSubtotal(NumberOfChecks.ToString() + " Checks Cleared for " + FormatAmount(statement.ChecksTotal));
            }

        }


        static void AddAdvances(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction advance, ref PdfPTable aTable)
        {

            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource advanceSource;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory advanceCategory;
            DateTime advanceDate = DateTime.Now;
            string advanceDescription = string.Empty;

            int advanceLength = advance.Items.Length;
            int counter = 0;
            string item;
            decimal advanceAmount = 0;
            //should be number of advances instead of payments
            int rowBreakPointIndex = (int)Math.Ceiling((double)NumberOfPayments / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("LOAN ADVANCES AND OTHER CHARGES");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            while (counter < advanceLength)
            {
                item = advance.ItemsElementName[counter].ToString();

                switch (item)
                {
                    case "source":
                        advanceSource = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource)advance.Items[counter];
                        sortLoanSource(advance, advanceSource, ref table);
                        break;
                    case "postingDate":
                        advanceDate = (DateTime)advance.Items[counter];
                        break;
                    case "category":
                        advanceCategory = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory)advance.Items[counter];
                        break;
                    case "grossAmount":
                        advanceAmount = (decimal)advance.Items[counter];
                        break;
                    case "description":
                        advanceDescription = (string)advance.Items[counter];
                        break;
                    case "transactionSerial":
                        break;
                    case "monetarySerial":
                        break;
                    case "principal":
                        break;
                    case "newBalance":
                        break;
                    default:
                        Console.WriteLine("ADVANCE ITEM {0}: {1}", item, advance.Items[counter]);
                        break;
                }
                /*
                for (int i = 0; i < advance.Advances.Count; i++)
                {
                    if ((i + 1) <= rowBreakPointIndex)
                    {
                        rows.Add(new SortTableRow());
                        rows[i].Column.Add(advanceDate.ToString("MMM dd"));
                        rows[i].Column.Add(FormatAmount(advanceAmount));
                        rows[i].Column.Add(advanceDescription);
                        rows[i].Column.Add(string.Empty);
                        rows[i].Column.Add(string.Empty);
                        rows[i].Column.Add(string.Empty);
                    }
                    else
                    {
                        rows[i - rowBreakPointIndex].Column[3] = advanceDate.ToString("MMM dd");
                        rows[i - rowBreakPointIndex].Column[4] = FormatAmount(advanceAmount);
                        rows[i - rowBreakPointIndex].Column[5] = advanceDescription;
                    }
                }
                */
                for (int i = 0; i < rows.Count; i++)
                {
                    AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                    AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                    AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                    AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                    AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                    AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                }

                counter++;
            }

            Doc.Add(table);



        }

        static void addRefinance(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction refinance, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource refinanceSource;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory refinanceCategory;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionTransferOption refinanceTransferOption;

            int refinanceLength = refinance.Items.Length;
            int counter = 0;
            string item;

            while (counter < refinanceLength)
            {
                item = refinance.ItemsElementName[counter].ToString();

                switch (item)
                {
                    case "source":
                        refinanceSource = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource)refinance.Items[counter];
                        sortLoanSource(refinance, refinanceSource, ref table);    
                        break;
                    case "category":
                        refinanceCategory = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory)refinance.Items[counter];
                        break;
                    case "transferOption":
                        refinanceTransferOption = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionTransferOption)refinance.Items[counter];
                        break;
                    default:
                        Console.WriteLine("REFINANCE ITEM {0}: {1}", item, refinance.Items[counter]);
                        break;
                }

                counter++;
            }
        }

        //loan transaction category: comment
        static void addLoanCommentTransaction(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction comment, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource loanCommentSource;
            statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory loanCommentCategory;
            string description = string.Empty;
            int loanCommentLength = comment.Items.Length;
            int counter = 0;
            string item;


            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            
            while (counter < loanCommentLength)
            {
                item = comment.ItemsElementName[counter].ToString();
           
                switch(item)
                {
                    case "source":
                        loanCommentSource = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionSource)comment.Items[counter];
                        break;
                    case "category":
                        loanCommentCategory = (statementProductionEnvelopeStatementAccountSubAccountLoanTransactionCategory)comment.Items[counter];
                        break;
                    case "description":
                        description = comment.Items[counter].ToString();
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(description, GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20;
                        cell.AddElement(p);
                        cell.PaddingTop = -4f;
                        cell.BorderWidth = 0f;
                        cell.NoWrap = true;
                        table.AddCell(cell);
                        break;
                    case "transactionSerial":
                        break;
                    case "monetarySerial":
                        break;
                    case "postingDate":
                        break;
                    case "newBalance":
                        break;
                    default:
                        Console.WriteLine("LOAN COMMENT ITEM: {0}", comment.ItemsElementName[counter]);
                        break;
                }
               
                counter++;
            }

            AddAccountTransactionAmount(0, ref table); // Additions
            AddAccountTransactionAmount(0, ref table); // Subtractions
            AddAccountTransactionAmount(0, ref table);
        }


        //share transaction category: comment
        static void addShareCommentTransaction(statementProductionEnvelopeStatementAccountSubAccountShareTransaction comment, ref PdfPTable table)
        {
            int shareCommentLength = comment.Items.Length;
            int counter = 0;
            string item;
            string description = string.Empty;

            while (counter < shareCommentLength)
            {
                item = comment.ItemsElementName[counter].ToString();

                switch(item)
                {
                    case "description":
                        description = comment.Items[counter].ToString();
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(description, GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20;
                        cell.AddElement(p);
                        cell.PaddingTop = -4f;
                        cell.BorderWidth = 0f;
                        cell.NoWrap = true;
                        table.AddCell(cell);
                        break;
                    case "transactionSerial":
                        break;
                    case "monetarySerial":
                        break;
                    case "postingDate":
                         PdfPCell Pcell = new PdfPCell();
                         Chunk Pchunk = new Chunk(comment.Items[counter].ToString(), GetNormalFont(9f));
                         Pchunk.SetCharacterSpacing(0f);
                         Paragraph par = new Paragraph(Pchunk);
                         Pcell.AddElement(par);
                         Pcell.PaddingTop = -4f;
                         Pcell.BorderWidth = 0f;
                         table.AddCell(Pcell);
                        break;
                    case "newBalance":
                        break;
                    case "category":
                        break;
                    default:
                        Console.WriteLine("SHARE COMMENT ITEM:{0}", item);
                        break;
                }
                counter++;
            }
            AddAccountTransactionAmount(0, ref table); // Additions
            AddAccountTransactionAmount(0, ref table); // Subtractions
            AddAccountTransactionAmount(0, ref table);
        }
    
        static void AddAddress(statementProductionEnvelope statement, ref PdfPTable table)
        {
            statementProductionEnvelopeAddress newAddress;
            statementProductionEnvelopeAddressCategory addressCategory;
            newAddress = statement.address;
            statementProductionEnvelopePerson currentPerson;
            PdfPCell cell = new PdfPCell();
            string fullAddress = statement.ToString();

            if(statement.address.category != null)
            {
                addressCategory = statement.address.category;
                fullAddress += addressCategory.Value;

            }

            Chunk chunk = new Chunk(fullAddress, GetNormalFont(9));

            if(statement.person != null)
            {
                currentPerson = statement.person;
                chunk.Append("\n" + currentPerson.firstName + currentPerson.middleName + currentPerson.lastName + currentPerson.suffix + ", " + currentPerson.type);
            }

            chunk.setLineHeight(9f);
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 66;
            cell.AddElement(p);
            cell.PaddingTop = 60f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            
        }

        public static void addHeaderBalances(statementProductionEnvelopeStatementAccountSubAccountShare account, statementProductionEnvelopeStatementAccountSubAccountLoan loan, string description, ref PdfPTable table)
        {

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk("Account  Balances  at  a  Glance:", GetBoldFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 81;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            PdfPTable balancesTable = new PdfPTable(2);
            float[] tableWidths = new float[] { 60f, 40f };
            balancesTable.SetWidthPercentage(tableWidths, Doc.PageSize);
            balancesTable.TotalWidth = 300f;
            balancesTable.LockedWidth = true;

            AddHeaderBalanceTitle(description, ref balancesTable);
            if (account != null)
            {
                AddHeaderBalanceValue(account.beginning.balance, ref balancesTable);
            }
            else if (loan != null)
            {
                AddHeaderBalanceValue(loan.beginning.balance, ref balancesTable);
            }

            cell.AddElement(balancesTable);
            table.AddCell(cell);
        }


        static void GetBalance(statementProduction statement, ref PdfPTable table)
        {
           

        }

        

        static void AddHeaderBalanceTitle(string title, ref PdfPTable balancesTable)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetNormalFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 75;
            cell.AddElement(p);
            cell.PaddingTop = -7f;
            cell.BorderWidth = 0f;
            balancesTable.AddCell(cell);
        }

        static void AddHeaderBalanceValue(decimal value, ref PdfPTable balancesTable)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(FormatAmount(value), GetBoldFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            p.IndentationRight = 28;
            cell.AddElement(p);
            cell.PaddingTop = -7f;
            cell.BorderWidth = 0f;
            balancesTable.AddCell(cell);
        }

        static void AddStatementHeading(string text, float indentationLeft, float paddingTop)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 612f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(text, GetNormalFont(12f));
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = indentationLeft;
            cell.AddElement(p);
            cell.PaddingTop = paddingTop;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddInvisibleAccountNumber(statementProduction statement)
        {
            for(int i = 0; i < statement.epilogue.envelopeCount; i++)
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 612f;
                table.LockedWidth = true;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(statement.envelope[0].statement.account.accountNumber.ToString(), GetNormalFont(5f));
                Paragraph p = new Paragraph(chunk);
                p.Font.SetColor(255, 255, 255);
                cell.AddElement(p);
                cell.PaddingTop = -9f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }
        }


       
        public static void addEndingLoanBalance(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction transaction, ref PdfPTable table)
        {
           
        }


    


        public static void sortAdditionalNames(statementProductionEnvelopeStatementAccountSubAccountSharePerson person)
        {
            string pCategory;

          
            if (person.personLinkCategory != null)
            {
                pCategory = (string)person.personLinkCategory.option;

                switch (pCategory)
                {
                    case "-":
                        break;
                    case "PR":
                        break;
                    case "TP":
                        addName(person, "Tax Plan Owner");
                        break;
                    case "JT":
                        addName(person, "Joint Owner");
                        break;
                    case "CB":
                        addName(person, "Co-borrower");
                        break;
                    case "AS":
                        addName(person, "Authorized signer");
                        break;
                    case "PA":
                        addName(person, "Power of attorney");
                        break;
                    case "TR":
                        addName(person, "Trustee");
                        break;
                    case "CU":
                        addName(person, "Custodian");
                        break;
                    case "BE":
                        addName(person, "Beneficiary");
                        break;
                    case "CS":
                        addName(person, "Co-signer");
                        break;
                    case "CO":
                        addName(person, "Collateral owner");
                        break;
                    case "SA":
                        addName(person, "Statement addressee");
                        break;
                    case "OT":
                        addName(person, "Other related party");
                        break;
                    default:
                        break;
                }

            }
        }

        public static void addName(statementProductionEnvelopeStatementAccountSubAccountSharePerson newName, string value)
        {
            statementProductionEnvelopeStatementAccountSubAccountSharePerson additionalName;
            additionalName = newName;

            if (additionalName.firstName != null)
            {
                //Console.WriteLine("Additional First Name: {0}", additionalName.firstName);
            }
            if (additionalName.serial > 0)
            {
                //Console.WriteLine("Additional Serial: {0}", additionalName.serial);
            }
            if (additionalName.middleName != null)
            {
                //Console.WriteLine("Additional Middle Name: {0}", additionalName.middleName);
            }
            if (additionalName.lastName != null)
            {
                //Console.WriteLine("Additional Last name: {0}", additionalName.lastName);
            }

            
        }




        public static void addLoanFees(statementProductionEnvelopeStatementAccountSubAccountLoanFeeTransaction feetransaction)
        {
            statementProductionEnvelopeStatementAccountSubAccountLoanFeeTransactionOpenEndLoanFeeIndicator openEnd;
            statementProductionEnvelopeStatementAccountSubAccountLoanFeeTransactionCategory feeCategory;
            statementProductionEnvelopeStatementAccountSubAccountLoanFeeTransactionSource feeSource;

            PdfPTable table = new PdfPTable(4);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 51, 234, 48, 192 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 0f;
            AddLoanTransactionTitle("Date", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Description", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Amount", Element.ALIGN_RIGHT, 1, ref table);
            AddLoanTransactionTitle(string.Empty, Element.ALIGN_RIGHT, 1, ref table);


            if (feetransaction.openEndLoanFeeIndicator != null)
            {
                openEnd = feetransaction.openEndLoanFeeIndicator;
                Console.WriteLine("Open End Indicator: {0}", feetransaction.openEndLoanFeeIndicator.Value);
            }

            if (feetransaction.transactionSerial > 0)
            {
                Console.WriteLine("Transaction serial: {0}", feetransaction.transactionSerial);
            }
            if (feetransaction.monetarySerial > 0)
            {
                Console.WriteLine("Monetary Serial: {0}", feetransaction.monetarySerial);
            }

            if (feetransaction.postingDate != null)
            {
                Console.WriteLine("Fee posting date: {0}", feetransaction.postingDate);
            }

            if (feetransaction.category != null)
            {
                feeCategory = feetransaction.category;
                Console.WriteLine("Fee Category: {0}", feeCategory.Value);
            }

            if (feetransaction.source != null)
            {
                feeSource = feetransaction.source;
                Console.WriteLine("Fee Source: {0}", feeSource.Value);
            }

            if (feetransaction.description != null)
            {
                Console.WriteLine("Fee Description: {0}", feetransaction.description);
            }

            if (feetransaction.grossAmount > 0)
            {
                Console.WriteLine("Gross Amount: {0}", feetransaction.grossAmount);
            }

            if (feetransaction.lateFee > 0)
            {
                Console.WriteLine("Late Fee: {0}", feetransaction.lateFee);
              
            }

            if (feetransaction.newBalance > 0)
            {
                Console.WriteLine("New balance: {0}", feetransaction.newBalance);
            }



            Doc.Add(table);
        
        }


        public static void NoTransactionsThisPeriodMessage(ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("No Transactions This Period", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            for (int i = 0; i < 5; i++)
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
            }
        }

        public static void addLoanInterest(statementProductionEnvelopeStatementAccountSubAccountLoanInterestCharge interest)
        {
            Console.WriteLine("Interest Posting Date: {0}", interest.postingDate);
            Console.WriteLine("Interest Charge: ${0}", interest.interest);
        }


        static void AddSectionHeading(string title)
        {
            if (Writer.GetVerticalPosition(false) <= 175)
            {
                Doc.NewPage();
            }

            AddHeadingStroke();
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(16f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            cell.AddElement(p);
            cell.PaddingTop = -7f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        
        static void AddBalanceForward(statementProduction account, ref PdfPTable table)
        {
            for (int i = 0; i < account.epilogue.accountCount; i++ )
            {
                // Adds Date
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(account.envelope[i].statement.account.openDate.ToString("MMM dd"), GetBoldItalicFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);

                }
                // Adds Transaction Description
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Balance Forward", GetBoldItalicFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                AddAccountTransactionAmount(0, ref table);
                AddAccountTransactionAmount(0, ref table);
                AddAccountBalance(account.envelope[i].statement.account.subAccount[i].share.beginning.balance, ref table);
            }
           
        }

        static void AddEndingBalance(statementProductionEnvelopeStatementAccountSubAccountShare account, ref PdfPTable table)
        {
            
                // Date
                 {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(account.endingStatementDate.ToString("MMM dd"), GetBoldItalicFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                // Transaction Description
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Ending Balance", GetBoldItalicFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

            AddAccountTransactionAmount(0, ref table);
            AddAccountTransactionAmount(0, ref table);
            AddAccountBalance(account.ending.balance, ref table);
            
        }

        static void AddShareClosed(statementProductionEnvelopeStatementAccountSubAccountShare account, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountShare closedShare;
            closedShare = account;

                // Date
                {
                    PdfPCell cell = new PdfPCell();
                    //display closed date or posting date?
                    Chunk chunk = new Chunk(string.Empty, GetBoldItalicFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);

                }
            

            // Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(closedShare.description + " " + "Closed\n*** This is the final statement you will receive for this account ***\n*** Please retain this final statement for tax reporting purposes ***", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.NoWrap = true;
                cell.PaddingTop = -1f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddAccountTransactionAmount(0, ref table);
            AddAccountTransactionAmount(0, ref table);
            AddAccountTransactionAmount(0, ref table);
        }

        static void AddLoanClosed(statementProductionEnvelopeStatementAccountSubAccountLoan loan, ref PdfPTable table)
        {
            statementProductionEnvelopeStatementAccountSubAccountLoan tLoan;
            tLoan = loan;

            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(tLoan.description + " " + "Closed\n*** This is the final statement you will receive for this account ***\n*** Please retain this final statement for tax reporting purposes ***", GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.NoWrap = true;
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            for (int i = 0; i < 5; i++)
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
            }
        }

        static void AddAccountTransactionAmount(decimal amount, ref PdfPTable table)
        {
            string amountFormatted = string.Empty;

            if (amount != 0)
            {
                amountFormatted = FormatAmount(amount);
            }

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddLoanAccountTransactionAmount(decimal amount, ref PdfPTable table)
        {
            string amountFormatted = FormatAmount(amount);

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetNormalFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddMoneyPerksTransactionAmount(int amount, ref PdfPTable table)
        {
            string amountFormatted = string.Empty;

            if (amount != 0)
            {
                amountFormatted = amount.ToString();
            }

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddAccountBalance(decimal balance, ref PdfPTable table)
        {
            string amountFormatted = FormatAmount(balance);

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddMoneyPerksBalance(int balance, ref PdfPTable table)
        {
            string amountFormatted =balance.ToString();

            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(amountFormatted, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);
        }

        static void AddAccountTransactionTitle(string title, int alignment, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
        }

        static void AddAccountSubHeading(string subtitle, bool stroke)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            float cellPaddingTop = -1f;

            if (stroke)
            {
                cellPaddingTop = -6f;
                AddSubHeadingStroke();
            }

            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(subtitle, GetBoldFont(12f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            cell.AddElement(p);
            cell.PaddingTop = cellPaddingTop;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }
        /*
        static void AddCheckHolds(statementProductionEnvelopeStatementAccount account)
        {
            foreach(CheckHold hold in account.CheckHolds)
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Check hold placed on " + hold.EffectiveDate.ToString("MM/dd/yyyy") + " in the amount of $" + FormatAmount(hold.Amount) + " to be released on " + hold.ExpiredDate.ToString("MM/dd/yyyy"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.IndentationLeft = 70;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }
        }
        */
        /*
        static void AddChecks(statementProductionEnvelopeStatementAccountSubAccountShare check)
        {

            statementProductionEnvelopeStatementAccountSubAccountShareTransaction memberCheck = new statementProductionEnvelopeStatementAccountSubAccountShareTransaction();
            DateTime checkDate = DateTime.Now;
            string checkDescription = string.Empty;
            decimal checkAmount = 0;

            bool asteriskFound = false;
            int rowBreakPointIndex = (int)Math.Ceiling((double)memberCheck.Checks.Count / 2);
            PdfPTable table = new PdfPTable(7);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 91, 43, 76, 105, 91, 43, 76 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;

            AddSortTableHeading("CHECK SUMMARY");

            AddSortTableTitle("Check #", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Date", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle(string.Empty, Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Check #", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Date", Element.ALIGN_RIGHT, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < memberCheck.Checks.Count; i++)
            {
                if (memberCheck.Checks[i].CheckNumber.Contains('*'))
                {
                    asteriskFound = true;
                }

                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    rows[i].Column.Add(memberCheck.Checks[i].CheckNumber);
                    rows[i].Column.Add(FormatAmount(checkAmount));
                    rows[i].Column.Add(checkDate.ToString("MMM dd"));
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    rows[i - rowBreakPointIndex].Column[4] = memberCheck.Checks[i].CheckNumber;
                    rows[i - rowBreakPointIndex].Column[5] = FormatAmount(checkAmount);
                    rows[i - rowBreakPointIndex].Column[6] = checkDate.ToString("MMM dd");
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Check #
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_RIGHT, ref table); // Adds Date
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_RIGHT, ref table); // Empty column title
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_LEFT, ref table);  // Adds Check #
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[6], Element.ALIGN_RIGHT, ref table); // Adds Date
            }

            Doc.Add(table);

            if (asteriskFound)
            {
                AddChecksFootnote();
            }

            if (memberCheck.Checks.Count() > 1)
            {
                AddSortTableSubtotal(memberCheck.Checks.Count().ToString() + " Checks Cleared for " + FormatAmount(memberCheck.ChecksTotal));
            }

            //memberCheck.Checks.Add(new Check(checkDescription, checkAmount, checkDate));
        }
       */
        /*
        
        static void AddAtmWithdrawals(statementProductionEnvelopeStatementAccountSubAccountShare account)
        {
            statementProductionEnvelopeStatementAccountSubAccountShareTransaction memberAtmWithdrawal = new statementProductionEnvelopeStatementAccountSubAccountShareTransaction();
            DateTime WithdrawalDate = DateTime.Now;
            string WithdrawalDescription = string.Empty;
            decimal WithdrawalAmount = 0;


            int rowBreakPointIndex = (int)Math.Ceiling((double)memberAtmWithdrawal.AtmWithdrawals.Count / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("ATM WITHDRAWALS AND OTHER CHARGES");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < memberAtmWithdrawal.AtmWithdrawals.Count; i++)
            {
                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    rows[i].Column.Add(WithdrawalDate.ToString("MMM dd"));
                    rows[i].Column.Add(FormatAmount(WithdrawalAmount));
                    rows[i].Column.Add(WithdrawalDescription);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    rows[i - rowBreakPointIndex].Column[3] = WithdrawalDate.ToString("MMM dd");
                    rows[i - rowBreakPointIndex].Column[4] = FormatAmount(WithdrawalAmount);
                    rows[i - rowBreakPointIndex].Column[5] = WithdrawalDescription;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);
            memberAtmWithdrawal.AtmWithdrawals.Add(new AtmWithdrawal(WithdrawalDescription, WithdrawalAmount, WithdrawalDate));
           

            if (memberAtmWithdrawal.AtmWithdrawals.Count() > 1)
            {
                AddSortTableSubtotal(memberAtmWithdrawal.AtmWithdrawals.Count().ToString() + " ATM Withdrawals and Other Charges for " + FormatAmount(memberAtmWithdrawal.AtmWithdrawalsTotal));
            }

            
        }
        
        */
        /*
        static void AddAtmDeposits(statementProductionEnvelopeStatementAccountSubAccountShare account)
        {
            statementProductionEnvelopeStatementAccountSubAccountShareTransaction memberAtmDeposit = new statementProductionEnvelopeStatementAccountSubAccountShareTransaction();
            DateTime AtmDepositDate = DateTime.Now;
            string AtmDepositDescription = string.Empty;
            decimal AtmDepositAmount = 0;

            int rowBreakPointIndex = (int)Math.Ceiling((double)memberAtmDeposit.AtmDeposits.Count / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("ATM DEPOSITS AND OTHER CHARGES");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < memberAtmDeposit.AtmDeposits.Count; i++)
            {
                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    rows[i].Column.Add(AtmDepositDate.ToString("MMM dd"));
                    rows[i].Column.Add(FormatAmount(AtmDepositAmount));
                    rows[i].Column.Add(AtmDepositDescription);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    rows[i - rowBreakPointIndex].Column[3] = AtmDepositDate.ToString("MMM dd");
                    rows[i - rowBreakPointIndex].Column[4] = FormatAmount(AtmDepositAmount);
                    rows[i - rowBreakPointIndex].Column[5] = AtmDepositDescription;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (memberAtmDeposit.AtmDeposits.Count() > 1)
            {
                AddSortTableSubtotal(memberAtmDeposit.AtmDeposits.Count().ToString() + " ATM Deposits and Other Charges for " + FormatAmount(memberAtmDeposit.AtmDepositsTotal));
            }
        }
        */

        //Share Transaction Category: WITHDRAWAL
        static void AddWithdrawal(statementProductionEnvelopeStatementAccountSubAccountShareTransaction withdrawal)
        {
            //statementProductionEnvelopeStatementAccountSubAccountShareTransaction memberAccount = new statementProductionEnvelopeStatementAccountSubAccountShareTransaction();
            
            
            NumberOfWithdrawals++;
            int withdrawalLength = withdrawal.Items.Length;
            int counter = 0;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionCategory withdrawalCategory;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource withdrawalSource;
            statementProductionEnvelopeStatementAccountSubAccountShareTransactionSubCategory withdrawalSubCategory;
            string item;
            int i = 0;
            decimal withdrawalAmount = 0;
            string withdrawalDescription = string.Empty;
            DateTime withdrawalDate = DateTime.Now;

            
            
            int rowBreakPointIndex = (int)Math.Ceiling((double)NumberOfDeposits / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;
            
            AddSortTableHeading("WITHDRAWALS AND OTHER CHARGES");
            
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            
            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();
            

            while (counter < withdrawalLength)
            {
                //if (i < NumberOfWithdrawals)
                //{
                    item = withdrawal.ItemsElementName[counter].ToString();
                    

                    switch (item)
                    {
                        case "category":
                            withdrawalCategory = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionCategory)withdrawal.Items[counter];
                            //Console.WriteLine("Withdrawal category {0}", withdrawalCategory.Value);
                            break;
                        case "source":
                            withdrawalSource = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionSource)withdrawal.Items[counter];
                            //Console.WriteLine("Withdrawal source {0}", withdrawalSource.Value);
                            break;
                        case "postingDate":
                            //Console.WriteLine("Withdrawal date:{0}", withdrawal.Items[counter]);
                            withdrawalDate = (DateTime)withdrawal.Items[counter];
                            break;
                        case "grossAmount":
                            withdrawalAmount = (decimal)withdrawal.Items[counter];
                            break;
                        case "description":
                            withdrawalDescription = (string)withdrawal.Items[counter];
                            break;
                        case "transactionSerial":
                            break;
                        case "monetarySerial":
                            break;
                        case "principal":
                            break;
                        case "newBalance":
                            break;
                        case "subCategory":
                            withdrawalSubCategory = (statementProductionEnvelopeStatementAccountSubAccountShareTransactionSubCategory)withdrawal.Items[counter];
                            AddWithdrawalSubCategory(withdrawalSubCategory);
                            break;
                        case "accountNumber":
                            break;
                        case "terminalLocation":
                            break;
                        case "terminalId":
                            break;
                        case "terminalCity":
                            break;
                        case "transferOption":
                            break;
                        case "transferId":
                            break;
                        case "transferIdCategory":
                            break;
                        case "terminalState":
                            break;
                        case "merchantName":
                            break;
                        case "merchantType":
                            break;
                        case "maskedCardNumber":
                            break;
                        case "transactionReference":
                            break;
                        case "transactionDate":
                            break;
                        case "draftNumber":
                            break;
                        case "draftTracer":
                            break;
                        case "routingNumber":
                            break;
                        case "transactionAmount":
                            break;
                        case "availableAmount":
                            break;
                        case "certificatePenalty":
                            break;
                        case "transferAccountNumber":
                            break;
                        case "transferName":
                            break;
                        case "transferIdDescription":
                            break;
                        case "adjustmentOption":
                            break;
                        case "feeClassification":
                            break;
                        default:
                            Console.WriteLine("WITHDRAWAL ITEM:{0}", item);
                            break;
                    }
/*
                    if ((i + 1) <= rowBreakPointIndex)
                    {
                        rows.Add(new SortTableRow());
                        rows[i].Column.Add(withdrawalDate.ToString("MMM dd"));
                        rows[i].Column.Add(FormatAmount(withdrawalAmount));
                        rows[i].Column.Add(withdrawalDescription);
                        rows[i].Column.Add(string.Empty);
                        rows[i].Column.Add(string.Empty);
                        rows[i].Column.Add(string.Empty);
                    }
                    else
                    {
                        rows[i - rowBreakPointIndex].Column[3] = withdrawalDate.ToString("MMM dd");
                        rows[i - rowBreakPointIndex].Column[4] = FormatAmount(withdrawalAmount);
                        rows[i - rowBreakPointIndex].Column[5] = withdrawalDescription;
                    }
                    */
                    counter++;
                    i++;

                    /*
                    for(int j = 0; j < rows.Count; j++)
                    {
                        AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                        AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                        AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                        AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                        AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                        AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                    }
                     */ 
               // }

               // memberAccount.Withdrawals.Add(new Withdrawal(withdrawalDescription, withdrawalAmount, withdrawalDate));
                Doc.Add(table);
                /*
                if (withdrawal.Withdrawals.Count() > 1)
                {
                    AddSortTableSubtotal(withdrawal.Withdrawals.Count().ToString() + " Withdrawals and Other Charges for " + FormatAmount(memberAccount.WithdrawalsTotal));
                }
                 */ 
            }

        }

        static void AddWithdrawalSubCategory(statementProductionEnvelopeStatementAccountSubAccountShareTransactionSubCategory subcategory)
        {

        }
        
 
        
        
        
         
        /*
        static void AddLoanPaymentsSortTable(Loan loan)
        {
            List<Deposit> loanPayments = loan.Payments;

            int rowBreakPointIndex = (int)Math.Ceiling((double)loanPayments.Count / 2);
            PdfPTable table = new PdfPTable(6);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 36, 63, 163.5f, 36, 63, 163.5f };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            AddSortTableHeading("LOAN PAYMENTS AND OTHER CREDITS");

            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);
            AddSortTableTitle("Date", Element.ALIGN_LEFT, ref table);
            AddSortTableTitle("Amount", Element.ALIGN_RIGHT, ref table);
            AddSortTableTitle("Description", Element.ALIGN_LEFT, 10f, ref table);

            // Create columns
            List<SortTableRow> rows = new List<SortTableRow>();

            for (int i = 0; i < loanPayments.Count; i++)
            {
                if ((i + 1) <= rowBreakPointIndex)
                {
                    rows.Add(new SortTableRow());
                    rows[i].Column.Add(loanPayments[i].Date.ToString("MMM dd"));
                    rows[i].Column.Add(FormatAmount(Math.Abs(loanPayments[i].Amount)));
                    rows[i].Column.Add(loanPayments[i].Description);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                    rows[i].Column.Add(string.Empty);
                }
                else
                {
                    rows[i - rowBreakPointIndex].Column[3] = loanPayments[i].Date.ToString("MMM dd");
                    rows[i - rowBreakPointIndex].Column[4] = FormatAmount(Math.Abs(loanPayments[i].Amount));
                    rows[i - rowBreakPointIndex].Column[5] = loanPayments[i].Description;
                }
            }

            for (int i = 0; i < rows.Count; i++)
            {
                AddSortTableValue(rows[i].Column[0], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[1], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[2], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
                AddSortTableValue(rows[i].Column[3], Element.ALIGN_LEFT, ref table);  // Adds Date
                AddSortTableValue(rows[i].Column[4], Element.ALIGN_RIGHT, ref table); // Adds Amount
                AddSortTableValue(rows[i].Column[5], Element.ALIGN_LEFT, 10f, ref table); // Adds Description
            }

            Doc.Add(table);

            if (loanPayments.Count() > 1)
            {
                AddSortTableSubtotal(loanPayments.Count().ToString() + " Payments and Other Credits for " + FormatAmount(Math.Abs(loan.PaymentsTotal)));
            }
        }
        */
        static void AddSortTableHeading(string title)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 10f;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_CENTER;
            cell.AddElement(p);
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddSortTableTitle(string title,  int alignment, ref PdfPTable table)
        {
            AddSortTableTitle(title, alignment, 0, ref table);
        }

        static void AddSortTableTitle(string title, int alignment, float indentation, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            p.IndentationLeft = indentation;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
        }

        static void AddSortTableValue(string value, int alignment, ref PdfPTable table)
        {
            AddSortTableValue(value, alignment, 0, ref table);
        }

        static void AddSortTableSubtotal(string value)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(value, GetBoldItalicFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 70;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddChecksFootnote()
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk("* Asterisk next to number indicates skip in number sequence", GetBoldItalicFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            //p.IndentationLeft = 70;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddSortTableValue(string value, int alignment, float indentation, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(value, GetNormalFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            p.IndentationLeft = indentation;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
        }

        static void AddTotalFees(statementProductionEnvelopeStatementAccountSubAccountLoanFeeTransaction fees)
        {
            PdfPTable table = new PdfPTable(5);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 12, 153, 79, 93, 188 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 10f;

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddFeeTitle("", new Border(1, 1, 0, 1) , ref table);
            AddFeeTitle("Total for\nthis period", new Border(1, 1, 0, 1), ref table);
            AddFeeTitle("Total\nyear-to-date", new Border(1, 1, 0, 1), ref table);

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Overdraft Fees
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Overdraft Fees", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 0;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }
            //Check for All fees (withdrawal fees, overdraft fees, etc...)
           // AddFeeValue(fees.lateFee.TotalOverdraftFee.AmountThisPeriod, new Border(1, 0, 0, 0), -2f, ref table);
           // AddFeeValue(fee.TotalOverdraftFee.AmountYtd, new Border(1, 0, 0, 0), -2f, ref table);

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -2f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -8f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Returned Item Fees
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Returned Item Fees", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -8f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                cell.BorderWidthTop = 0;
                cell.BorderWidthRight = 0;
                cell.BorderWidthBottom = 1;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }
            //total returned item fees this period/ytd
            //AddFeeValue(fee.TotalReturnedItemFee.AmountThisPeriod, new Border(1, 0, 0, 1), -8f, ref table);
            //AddFeeValue(fee.TotalReturnedItemFee.AmountYtd, new Border(1, 0, 0, 1), -8f, ref table);

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 6f;
                cell.PaddingTop = -8f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                table.AddCell(cell);
            }

            Doc.Add(table);
        }

        static void AddFeeTitle(string title, Border border, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(11f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_CENTER;
            cell.AddElement(p);
            cell.Padding = 6f;
            cell.PaddingTop = -2f;
            cell.BorderWidth = 0f;
            cell.BorderWidthLeft = border.WidthLeft;
            cell.BorderWidthTop = border.WidthTop;
            cell.BorderWidthRight = border.WidthRight;
            cell.BorderWidthBottom = border.WidthBottom;
            cell.BorderColor = BaseColor.BLACK;
            table.AddCell(cell);
        }

        static void AddFeeValue(decimal value, Border border, float paddingTop, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(FormatAmount(value), GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_RIGHT;
            cell.AddElement(p);
            cell.Padding = 6f;
            cell.PaddingTop = paddingTop;
            cell.PaddingRight = 35f;
            cell.BorderWidth = 0f;
            cell.BorderWidthLeft = border.WidthLeft;
            cell.BorderWidthTop = border.WidthTop;
            cell.BorderWidthRight = border.WidthRight;
            cell.BorderWidthBottom = border.WidthBottom;
            cell.BorderColor = BaseColor.BLACK;
            table.AddCell(cell);
        }
       /* 
        static void AddLoanPaymentInformation(statementProductionEnvelopeStatementAccountSubAccountLoanTransaction payment)
        {
            // Annual percentage rate
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = null;
                if (loan.CreditLimit == null)
                {
                    chunk = new Chunk("Annual Percentage Rate:  " + loan.AnnualPercentageRate.ToString("N3") + "%", GetBoldFont(9f));
                }
                else
                {
                    chunk = new Chunk("Annual Percentage Rate:  " + loan.AnnualPercentageRate.ToString("N3") + "%    Credit Limit:    " + FormatAmount(loan.CreditLimit.Limit) + "    Available Credit:    " + FormatAmount(loan.CreditLimit.Available), GetBoldFont(9f));
                }
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.IndentationLeft = 28f;
                cell.AddElement(p);
                cell.PaddingTop = 7f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            // PAYMENT INFORMATION
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("PAYMENT INFORMATION", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = 7f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            // Summary table
            {
                PdfPTable table = new PdfPTable(3);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 93, 55, 381 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0f;

                // Previous Balance Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Previous Balance:", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // Previous Balance
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(loan.PreviousBalance.ToString("N"), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // New Balance Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("New Balance:", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // New Balance
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(loan.NewBalance.ToString("N"), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // Minimum Payment Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Minimum Payment:", GetBoldFont(9f));
                    if (loan.MinimumPayment == 0)
                    {
                        chunk = new Chunk("Minimum Payment: No Payment Due", GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    cell.NoWrap = true;
                    table.AddCell(cell);
                }
                // Minimum Payment
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    if (loan.MinimumPayment != 0)
                    {
                        chunk = new Chunk(loan.MinimumPayment.ToString("N"), GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // Payment Due Date Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Payment Due Date:", GetBoldFont(9f));
                    if (loan.PaymentDueDate.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        chunk = new Chunk("Payment Due Date: No Payment Due", GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    cell.NoWrap = true;
                    table.AddCell(cell);
                }
                // Payment Due Date
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    if (loan.PaymentDueDate.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        chunk = new Chunk(loan.PaymentDueDate.ToString("MM/dd/yyyy"), GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }

            // Next Payment Due Date after statement
            {
                PdfPTable table = new PdfPTable(3);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 180, 55, 290 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0;

                // Next Payment Due Date after statement Title
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Next Payment Due Date after statement:", GetBoldFont(9f));
                    if (loan.NextPaymentDueDateAfterStatement.ToString("MM/dd/yyyy") == "01/01/0001")
                    {
                        chunk = new Chunk("Next Payment Due Date after statement: No Payment Due", GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    cell.NoWrap = true;
                    table.AddCell(cell);
                }
                // Next Payment Due Date after statement
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    if (loan.NextPaymentDueDateAfterStatement.ToString("MM/dd/yyyy") != "01/01/0001")
                    {
                        chunk = new Chunk(loan.NextPaymentDueDateAfterStatement.ToString("MM/dd/yyyy"), GetBoldFont(9f));
                    }
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }
        }
        */ 
    


        static void AddSeeFeeSummaryMessage(statementProductionEnvelopeStatementAccountSubAccountLoan loan, ref PdfPTable table)
        {
            
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(loan.endingStatementDate.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("See Fee Summary Below", GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            for (int i = 0; i < 5; i++)
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
            }
        }

        static void AddLoanTransactionTitle(string title, int alignment, int numOfLines, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(10f);
            //chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.GRAY;
            cell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
            cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
            cell.AddElement(Underline(chunk, alignment, numOfLines));
            table.AddCell(cell);
        }

        static void AddLoanTransactionsFooter(string value)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 12f;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(value, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(11f);
            Paragraph p = new Paragraph(chunk);
            p.IndentationLeft = 15;
            cell.AddElement(p);
            cell.PaddingTop = -4f;
            cell.BorderWidth = 0f;
            table.AddCell(cell);
            Doc.Add(table);
        }

        static void AddFeeSummary(statementProductionEnvelopeStatementAccountSubAccountLoan loan)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            // Title
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                table.SpacingBefore = 12f;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("FEE SUMMARY", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                //p.IndentationLeft = 15;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.PaddingLeft = 10f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            AddLoanFees(loan);

            // TOTAL FEES FOR THIS PERIOD
            {
                PdfPTable table = new PdfPTable(4);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 51, 234, 48, 192 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0f;

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // TOTAL FEES FOR THIS PERIOD
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("TOTAL FEES FOR THIS PERIOD", GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_LEFT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // value
                {
                    PdfPCell cell = new PdfPCell();
                    if(loan.loanFeesChargedPeriodSpecified)
                    {
                        Chunk chunk = new Chunk(loan.loanFeesChargedPeriod.ToString("N"), GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -4f;
                        cell.BorderWidth = 0f;
                        cell.BorderColor = BaseColor.LIGHT_GRAY;
                        table.AddCell(cell);
                    }
                    else
                    {
                        Chunk chunk = new Chunk("No fees this period.", GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -4f;
                        cell.BorderWidth = 0f;
                        cell.BorderColor = BaseColor.LIGHT_GRAY;
                        table.AddCell(cell);
                    }
                   
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }
        }

        static void AddLoanFees(statementProductionEnvelopeStatementAccountSubAccountLoan loan)
        {
            statementProductionEnvelopeStatementAccountSubAccountLoanFeeTransaction fees;

            PdfPTable table = new PdfPTable(4);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 51, 234, 48, 192 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 0f;
            AddLoanTransactionTitle("Date", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Description", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Amount", Element.ALIGN_RIGHT, 1, ref table);
            AddLoanTransactionTitle(string.Empty, Element.ALIGN_RIGHT, 1, ref table);

            if(loan.feeTransaction != null)
            {
                fees = loan.feeTransaction;
                AddLoanFee(fees, ref table);
            }
            
            

            Doc.Add(table);
        }


        static void AddLoanFee(statementProductionEnvelopeStatementAccountSubAccountLoanFeeTransaction fee, ref PdfPTable table)
        {
 
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(fee.postingDate.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                //string description = fee.Description.Length > 30 ? fee.Description.Substring(0, 30) : fee.Description;
                string description = fee.description;

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddLoanAccountTransactionAmount(fee.grossAmount, ref table); // Amount

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -6f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddAccountBalance(fee.newBalance, ref table);
        }



        static void AddInterestChargedSummary(statementProductionEnvelopeStatementAccountSubAccountLoan loan)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            // Title
            {
                PdfPTable table = new PdfPTable(1);
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                table.SpacingBefore = 12f;
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("INTEREST CHARGED SUMMARY", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                //p.IndentationLeft = 15;
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.PaddingLeft = 10f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
                Doc.Add(table);
            }

            AddLoanInterestTransactions(loan);

            // TOTAL INTEREST FOR THIS PERIOD
            {
                PdfPTable table = new PdfPTable(4);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 51, 234, 48, 192 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                table.SpacingBefore = 0f;

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // TOTAL INTEREST FOR THIS PERIOD
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("TOTAL INTEREST FOR THIS PERIOD", GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_LEFT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                // value
                {
                    PdfPCell cell = new PdfPCell();
                    if(loan.interestChargedPeriodSpecified == true)
                    {
                        Chunk chunk = new Chunk(loan.interestChargedPeriod.ToString("N"), GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -4f;
                        cell.BorderWidth = 0f;
                        cell.BorderColor = BaseColor.LIGHT_GRAY;
                        table.AddCell(cell);
                    }
                    else
                    {
                        Chunk chunk = new Chunk("No interest fees charged this period.", GetNormalFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -4f;
                        cell.BorderWidth = 0f;
                        cell.BorderColor = BaseColor.LIGHT_GRAY;
                        table.AddCell(cell);
                    }
                   
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderColor = BaseColor.LIGHT_GRAY;
                    table.AddCell(cell);
                }

                Doc.Add(table);
            }
        }


        
        static void AddLoanInterestTransactions(statementProductionEnvelopeStatementAccountSubAccountLoan loan)
        {
            PdfPTable table = new PdfPTable(4);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 51, 234, 48, 192 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 0f;
            AddLoanTransactionTitle("Date", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Description", Element.ALIGN_LEFT, 1, ref table);
            AddLoanTransactionTitle("Amount", Element.ALIGN_RIGHT, 1, ref table);
            AddLoanTransactionTitle(string.Empty, Element.ALIGN_RIGHT, 1, ref table);
            /*
            foreach (statementProductionEnvelopeStatementAccountSubAccountLoan transaction in loan.transaction[])
            {
                if (transaction.InterestCharged > 0)
                {
                    AddLoanInterestTransaction(transaction, ref table);
                }
            }
             */ 

            Doc.Add(table);
        }
        /*
        static void AddLoanInterestTransaction(statementProduction transaction, ref PdfPTable table)
        {
            for(int i = 0; i < transaction.epilogue.loanCount; i++)
            {
                // Adds Date
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(transaction.envelope[i].statement.account.subAccount[i].share.transaction[i].Items[2].ToString("MMM dd"), GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                // Adds Transaction Description
                {
                    string description = transaction.DescriptionLine1.Length > 30 ? transaction.DescriptionLine1.Substring(0, 30) : transaction.DescriptionLine1;

                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(description, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                AddLoanAccountTransactionAmount(transaction.InterestCharged, ref table); // Amount

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.PaddingTop = -6f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
            }
            //AddAccountBalance(transaction.Balance, ref table);
        }
        */
        static void AddYearToDateTotals(statementProductionEnvelopeStatementAccountSubAccountLoan loan)
        {
            PdfPTable table = new PdfPTable(5);
            table.HeaderRows = 1;
            float[] tableWidths = new float[] { 12, 153, 79, 93, 188 };
            table.TotalWidth = 525f;
            table.SetWidths(tableWidths);
            table.LockedWidth = true;
            table.SpacingBefore = 20f;
            table.KeepTogether = true;

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Title
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("YEAR TO DATE TOTALS", GetBoldFont(9));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_LEFT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.PaddingLeft = 6f;
                cell.BorderWidth = 0f;
                cell.BorderWidthTop = 1f;
                cell.BorderWidthLeft = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldFont(9));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_LEFT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.BorderWidth = 0f;
                cell.BorderWidthTop = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldFont(9));
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(11f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_LEFT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = 0.5f;
                cell.BorderWidth = 0f;
                cell.BorderWidthTop = 1f;
                cell.BorderWidthRight = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Fees Charged this Year Title
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Fees Charged this Year", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingLeft = 6f;
                cell.BorderWidth = 0f;
                cell.BorderWidthLeft = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // Total Fees Charged this Year Value
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(FormatAmount(loan.loanFeesChargedYTD), GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.PaddingRight = 35f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                cell.BorderWidthRight = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            // Total Interest Charged this Year Title
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Total Interest Charged this Year", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                //cell.PaddingBottom = (loan.ExistedLastYear) ? 0f : 5f;
                cell.PaddingLeft = 6f;
                cell.BorderWidth = 0f;
                //cell.BorderWidthBottom = (loan.ExistedLastYear) ? 0f : 1f;
                cell.BorderWidthLeft = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // Total Interest Charged this Year Value
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(FormatAmount(loan.interestChargedYTD), GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_RIGHT;
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                //cell.PaddingBottom = (loan.ExistedLastYear) ? 0f : 5f;
                cell.PaddingRight = 35f;
                cell.BorderWidth = 0f;
                //cell.BorderWidthBottom = (loan.ExistedLastYear) ? 0f : 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                //cell.PaddingBottom = (loan.ExistedLastYear) ? 0f : 5f;
                cell.BorderWidth = 0f;
                //cell.BorderWidthBottom = (loan.ExistedLastYear) ? 0f : 1f;
                cell.BorderWidthRight = 1f;
                cell.BorderColor = BaseColor.BLACK;
                table.AddCell(cell);
            }

            // For layout only
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("", GetBoldFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.Padding = 0;
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            if (loan.interestLastYearSpecified)
            {
                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                // Total Fees Charged Last Year
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Total Fees Charged Last Year", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingLeft = 6f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 0f;
                    cell.BorderWidthLeft = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // Total Interest Charged this Year Value
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(FormatAmount(loan.interestChargedYTD), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingRight = 35f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 0f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 0f;
                    cell.BorderWidthRight = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }

                // Total Interest Charged Last Year
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("Total Interest Charged Last Year", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingLeft = 6f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 1f;
                    cell.BorderWidthLeft = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // Total Interest Charged Last Year Value
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk(FormatAmount(loan.interestChargedLastYear), GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    p.Alignment = Element.ALIGN_RIGHT;
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.PaddingRight = 35f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    cell.BorderWidthBottom = 1f;
                    cell.BorderWidthRight = 1f;
                    cell.BorderColor = BaseColor.BLACK;
                    table.AddCell(cell);
                }

                // For layout only
                {
                    PdfPCell cell = new PdfPCell();
                    Chunk chunk = new Chunk("", GetBoldFont(9f));
                    chunk.SetCharacterSpacing(0f);
                    Paragraph p = new Paragraph(chunk);
                    cell.AddElement(p);
                    cell.Padding = 0;
                    cell.PaddingTop = -4f;
                    cell.PaddingBottom = 5f;
                    cell.BorderWidth = 0f;
                    table.AddCell(cell);
                }
            }

            

            Doc.Add(table);
        }
        
        static void AddYtdSummaries(int loanCount, int accountCount, statementProductionEnvelopeStatementAccountSubAccount statement)
        {
            statementProductionEnvelopeStatementAccountSubAccountLoan loan;
            statementProductionEnvelopeStatementAccountSubAccountShare share;            
            PdfPTable leftTable;
            PdfPCell leftTableCell;
            decimal interestChargedYTD = 0;

            decimal loanFeesChargedYTD = 0;
            decimal certificatePenaltyYTD = 0;
            decimal dividendYTD = 0;

            AddSectionHeading("YTD SUMMARIES");


            // A table to create 2 columns
            {
                PdfPTable table = new PdfPTable(2);
                table.HeaderRows = 0;
                float[] tableWidths = new float[] { 255, 270 };
                table.TotalWidth = 525f;
                table.LockedWidth = true;
                table.SpacingBefore = 15f;

                // TOTAL DIVIDENDS PAID
                {
                    leftTable = new PdfPTable(2);
                    leftTable.HeaderRows = 0;
                    float[] leftTableWidths = new float[] { 214.5f, 48f };
                    leftTable.TotalWidth = 262.5f;
                    leftTable.SetWidths(leftTableWidths);
                    leftTable.LockedWidth = true;
                    leftTable.SpacingBefore = 0f;

                    // Title
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("TOTAL DIVIDENDS PAID", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // For layout only
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }
                    /*
                    foreach (Account account in statement.Accounts.OrderBy(o => o.Description).ToList())
                    {
                        // Account
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(account.Description, GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.IndentationLeft = 20f;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }

                        // Value
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(FormatAmount(account.Dividends), GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.Alignment = Element.ALIGN_RIGHT;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }
                    }
                    */
                    
                    // Total
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));

                        if (statement.share != null)
                        {
                            share = statement.share;
                            if (share.dividendYTDSpecified)
                            {
                                chunk = new Chunk("Total Year To Date Dividends Paid", GetNormalFont(9f));
                                dividendYTD = share.dividendYTD;
                            }
                            if (share.certificatePenaltyYTDSpecified)
                            {
                                certificatePenaltyYTD = share.certificatePenaltyYTD;
                            }

                        }

                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));

                        chunk = new Chunk(FormatAmount(dividendYTD), GetNormalFont(9f));
                       
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }


                        // Total
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                            chunk = new Chunk("Certificate Penalty", GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.IndentationLeft = 20f;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }

                        // Value
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                            chunk = new Chunk(FormatAmount(certificatePenaltyYTD), GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.Alignment = Element.ALIGN_RIGHT;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }
                   
                    // Total
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));



                        if (statement.loan != null)
                        {
                            loan = statement.loan;
                                interestChargedYTD = loan.interestChargedYTD;
                                chunk = new Chunk("Total Year To Date Interest Paid", GetNormalFont(9f));
                                loanFeesChargedYTD = loan.loanFeesChargedYTD;
       
                        }

                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // Value
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk(string.Empty, GetNormalFont(9f));
                        
                        chunk = new Chunk(FormatAmount(interestChargedYTD), GetNormalFont(9f));
                        
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.Alignment = Element.ALIGN_RIGHT;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    leftTableCell = new PdfPCell();
                    if(accountCount > 0) leftTableCell.AddElement(leftTable);
                    leftTableCell.BorderWidth = 0;
                    leftTableCell.Padding = 0;

                    table.AddCell(leftTableCell);
                }

                // TOTAL LOAN INTEREST PAID
                {
                    leftTable = new PdfPTable(2);
                    leftTable.HeaderRows = 0;
                    float[] leftTableWidths = new float[] { 214.5f, 48f };
                    leftTable.TotalWidth = 262.5f;
                    leftTable.SetWidths(leftTableWidths);
                    leftTable.LockedWidth = true;
                    leftTable.SpacingBefore = 0f;

                    // Title
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("TOTAL LOAN INTEREST PAID", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        p.IndentationLeft = 20f;
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }

                    // For layout only
                    {
                        PdfPCell cell = new PdfPCell();
                        Chunk chunk = new Chunk("", GetBoldFont(9f));
                        chunk.SetCharacterSpacing(0f);
                        Paragraph p = new Paragraph(chunk);
                        cell.AddElement(p);
                        cell.PaddingTop = -6f;
                        cell.BorderWidth = 0f;
                        leftTable.AddCell(cell);
                    }


                        // Loan
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk("Loan", GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.IndentationLeft = 20f;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }

                        // Value
                        {
                            PdfPCell cell = new PdfPCell();
                            Chunk chunk = new Chunk(FormatAmount(interestChargedYTD), GetNormalFont(9f));
                            chunk.SetCharacterSpacing(0f);
                            Paragraph p = new Paragraph(chunk);
                            p.Alignment = Element.ALIGN_RIGHT;
                            cell.AddElement(p);
                            cell.PaddingTop = -6f;
                            cell.BorderWidth = 0f;
                            leftTable.AddCell(cell);
                        }
                    

                    leftTableCell = new PdfPCell();
                    if(loanCount > 0) leftTableCell.AddElement(leftTable);
                    leftTableCell.BorderWidth = 0;
                    leftTableCell.Padding = 0;

                    table.AddCell(leftTableCell);
                }

 

                Doc.Add(table);
            }
        }
        
        static void AddMoneyPerksSummary(MoneyPerksStatement moneyPerksStatement)
        {
            if (moneyPerksStatement!=null)
            {
                AddSectionHeading("MONEYPERKS POINTS SUMMARY");

                PdfPTable table = new PdfPTable(5);
                table.HeaderRows = 1;
                float[] tableWidths = new float[] { 51, 280, 62, 65, 67 };
                table.TotalWidth = 525f;
                table.SetWidths(tableWidths);
                table.LockedWidth = true;
                AddMoneyPerksTransactionTitle("Date", Element.ALIGN_LEFT, 1, ref table);
                AddMoneyPerksTransactionTitle("Transaction Description", Element.ALIGN_LEFT, 1, ref table);
                AddMoneyPerksTransactionTitle("Points\nAwarded", Element.ALIGN_RIGHT, 2,  ref table);
                AddMoneyPerksTransactionTitle("Points\nRedeemed", Element.ALIGN_RIGHT, 2, ref table);
                AddMoneyPerksTransactionTitle("Balance", Element.ALIGN_RIGHT, 1,  ref table);
                AddMoneyPerksBeginningBalance(moneyPerksStatement, ref table);

                foreach (MoneyPerksTransaction transaction in moneyPerksStatement.Transactions)
                {
                    AddMoneyPerksTransaction(transaction, ref table);
                }

                AddMoneyPerksEndingBalance(moneyPerksStatement, ref table);
                Doc.Add(table);
            }
        }

        static void AddMoneyPerksTransactionTitle(string title, int alignment, int numOfLines, ref PdfPTable table)
        {
            PdfPCell cell = new PdfPCell();
            Chunk chunk = new Chunk(title, GetBoldFont(9f));
            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(10f);
            chunk.SetUnderline(0.75f, -2);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = alignment;
            cell.AddElement(p);
            cell.PaddingTop = 6f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.GRAY;
            cell.VerticalAlignment = PdfPCell.ALIGN_BOTTOM;
            cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
            cell.AddElement(Underline(chunk, alignment, numOfLines));
            table.AddCell(cell);
        }

        static void AddMoneyPerksBeginningBalance(MoneyPerksStatement moneyPerksStatement, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }
            // Adds Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Beginning Balance", GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksBalance(moneyPerksStatement.BeginningBalance, ref table);
        }

        static void AddMoneyPerksEndingBalance(MoneyPerksStatement moneyPerksStatement, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(string.Empty, GetBoldItalicFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);

            }
            // Adds Transaction Description
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk("Ending Balance", GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksTransactionAmount(0, ref table);
            AddMoneyPerksBalance(moneyPerksStatement.EndingBalance, ref table);
        }

        static void AddMoneyPerksTransaction(MoneyPerksTransaction transaction, ref PdfPTable table)
        {
            // Adds Date
            {
                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(transaction.Date.ToString("MMM dd"), GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }
            // Adds Transaction Description
            {
                string description = transaction.Description;

                PdfPCell cell = new PdfPCell();
                Chunk chunk = new Chunk(description, GetNormalFont(9f));
                chunk.SetCharacterSpacing(0f);
                Paragraph p = new Paragraph(chunk);
                cell.AddElement(p);
                cell.PaddingTop = -4f;
                cell.BorderWidth = 0f;
                table.AddCell(cell);
            }

            if (transaction.Amount >= 0)
            {
                AddMoneyPerksTransactionAmount(transaction.Amount, ref table); // Additions
                AddMoneyPerksTransactionAmount(0, ref table); // Subtractions
            }
            else
            {
                AddMoneyPerksTransactionAmount(0, ref table); // Additions
                AddMoneyPerksTransactionAmount(transaction.Amount, ref table); // Subtractions
            }

            AddMoneyPerksBalance(transaction.Balance, ref table);
        }
        
        static void AddTopAdvertising(statementProduction statement)
        {
            // Advertisement Bottom
            Font font = new Font(Font.FontFamily.HELVETICA, 12f, Font.BOLD, new BaseColor(0, 0, 0));
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 34f;
            PdfPCell cell = new PdfPCell();
            Chunk chunk = null;

            for (int i = 0; i < statement.AdvertisementTop.TotalLines; i++)
            {
                if (chunk == null)
                {
                    chunk = new Chunk(statement.AdvertisementTop.MessageLines[i], font);
                }
                else
                {
                    chunk.Append("\n" + statement.AdvertisementTop.MessageLines[i]);
                }
            }

            if (chunk == null)
            {
                chunk = new Chunk(string.Empty, font);
                table.SpacingBefore = 14f;
            }

            chunk.SetCharacterSpacing(0f);
            chunk.setLineHeight(12f);
            Paragraph p = new Paragraph(chunk);
            p.Alignment = Element.ALIGN_CENTER;
            p.IndentationLeft = 385;
            cell.AddElement(p);
            cell.PaddingTop = -1f;
            cell.BorderWidth = 0f;
            cell.BorderColor = BaseColor.LIGHT_GRAY;
            table.AddCell(cell);

            Doc.Add(table);
        }
        
        
        static void AddBottomAdvertising(statementProduction statement)
        {
            if (Writer.GetVerticalPosition(false) <= 130)
            {
                Doc.NewPage();
            }

            // Advertisement Bottom Stroke
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 10f;
            PdfPCell cell = new PdfPCell();
            cell.BorderWidthBottom = 5f;
            cell.BorderColor = BaseColor.BLACK;
            table.AddCell(cell);
            Doc.Add(table);

            // Advertisement Bottom
            Font font = new Font(Font.FontFamily.HELVETICA, 12f, Font.BOLD, new BaseColor(0, 0, 0));
            table = new PdfPTable(1);
            table.TotalWidth = 525f;
            table.LockedWidth = true;
            table.SpacingBefore = 10f;
            cell = new PdfPCell();
            Chunk chunk = null;

            for (int i = 0; i < statement.AdvertisementBottom.TotalLines; i++)
            {
                if (chunk == null)
                {
                    chunk = new Chunk(statement.AdvertisementBottom.MessageLines[i], font);
                }
                else
                {
                    chunk.Append("\n" + statement.AdvertisementBottom.MessageLines[i]);
                }
            }

            if (chunk != null)
            {
                chunk.SetCharacterSpacing(0f);
                chunk.setLineHeight(12f);
                Paragraph p = new Paragraph(chunk);
                p.Alignment = Element.ALIGN_CENTER;
                //p.IndentationLeft = 385;
                cell.AddElement(p);
                //cell.PaddingTop = -1f;
                cell.BorderWidth = 0f;
                cell.BorderColor = BaseColor.LIGHT_GRAY;
                table.AddCell(cell);
                Doc.Add(table);
            }
        }
        
        
        static void AddPageNumbersAndDisclosures(statementProduction statement)
        {
            DateTime beginDate = statement.envelope[0].statement.beginningStatementDate;
            DateTime endDate = statement.envelope[0].statement.endingStatementDate;
            decimal firstAccNumber = statement.envelope[0].statement.account.accountNumber;

            // Adds page numbers
            PdfReader statementReader = new PdfReader("C:\\" + TEMP_FILE_NAME);
            PdfReader statementBackReader = new PdfReader(Configuration.GetStatementDisclosuresTemplateFilePath());

            using (FileStream fs = new FileStream(Configuration.GetStatementsOutputPath() + statement.envelope[0].statement.account.accountNumber.ToString() + ".pdf", FileMode.Create, FileAccess.Write, FileShare.None))
            {
                using (PdfStamper stamper = new PdfStamper(statementReader, fs))
                {
                    stamper.SetFullCompression();
                    int pageCount = statementReader.NumberOfPages + 1; // Adds 1 for the disclosures page that will be added later
                    for (int i = 1; i <= pageCount - 1; i++)
                    {
                        if (i == 1)
                        {
                            // Page count on first page
                            Chunk chunk = new Chunk("Page:   1 of   " + pageCount, GetBoldFont(12f));
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(chunk), 578, 595, 0);
                            if (i != pageCount)
                            {
                                chunk = new Chunk("--- Continued on following page ---", GetBoldFont(9f));
                                ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_CENTER, new Phrase(chunk), 300, 20, 0);
                            }
                        }
                        else if (i != pageCount)
                        {
                            float startY = 750f;
                            float lineHeight = 10;

                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk(beginDate.ToString("MMM dd, yyyy") + "  thru  " + endDate.ToString("MMM dd, yyyy"), GetBoldFont(9f))), 578, startY, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Account  Number:   ******" + firstAccNumber.ToString(), GetBoldFont(9f))), 578, startY - lineHeight, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Page:  " + i.ToString() + " of " + pageCount.ToString(), GetBoldFont(9f))), 578, startY - (lineHeight * 2), 0);
                            Chunk chunk = new Chunk("--- Continued on reverse side ---", GetBoldFont(9f));
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_CENTER, new Phrase(chunk), 300, 20, 0);
                        }
                        else
                        {
                            float startY = 750f;
                            float lineHeight = 10;
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk(beginDate.ToString("MMM dd, yyyy") + "  thru  " + endDate.ToString("MMM dd, yyyy"), GetBoldFont(9f))), 578, startY, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Account  Number:   ******" + firstAccNumber.ToString(), GetBoldFont(9f))), 578, startY - lineHeight, 0);
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(new Chunk("Page:  " + i.ToString() + " of " + pageCount.ToString(), GetBoldFont(9f))), 578, startY - (lineHeight * 2), 0);
                        }
                    }

                    stamper.InsertPage(pageCount, PageSize.LETTER);
                    PdfContentByte cb = stamper.GetOverContent(pageCount);
                    PdfImportedPage p = stamper.GetImportedPage(statementBackReader, 1);
                    cb.AddTemplate(p, 0, 0);
                }
            }
        }
        
        static void AddHeadingStroke()
        {
            Doc.Add(Stroke(525f, 20f, 0, 5f, BaseColor.BLACK, Element.ALIGN_CENTER));
        }

        static void AddSubHeadingStroke()
        {
            Doc.Add(Stroke(525f, 10f, 0, 0.5f, BaseColor.BLACK, Element.ALIGN_CENTER));
        }


        static PdfPTable Stroke(float width, float spacingAbove, float spacingLeft, float thickness, BaseColor color, int alignment)
        {
            PdfPTable table = new PdfPTable(1);
            table.TotalWidth = width;
            table.LockedWidth = true;
            table.SpacingBefore = spacingAbove;
            PdfPCell cell = new PdfPCell();
            cell.BorderWidth = 0;
            cell.BorderWidthBottom = thickness;
            cell.BorderColor = color;
            cell.PaddingLeft = spacingLeft;
            table.AddCell(cell);
            table.HorizontalAlignment = alignment;
            return table;
        }

        /// <summary>
        /// Produces a stroke with a width that will fit underneath a chunk of text, even if the text is multiple lines long
        /// </summary>
        /// <param name="?"></param>
        /// <returns></returns>
        static PdfPTable Underline(Chunk textChunk, int alignment, int numOfLines)
        {
            if (numOfLines > 1)
            {
                string[] words = textChunk.ToString().Split('\n');
                string longestWord = string.Empty;
                Chunk wordChunk = null;
                Chunk longestWordChunk = new Chunk(longestWord);

                foreach (string word in words)
                {
                    wordChunk = new Chunk(word);
                    wordChunk.Font = textChunk.Font;
                    wordChunk.SetCharacterSpacing(textChunk.GetCharacterSpacing());

                    longestWordChunk = new Chunk(longestWord);
                    longestWordChunk.Font = textChunk.Font;
                    longestWordChunk.SetCharacterSpacing(textChunk.GetCharacterSpacing());

                    if (wordChunk.GetWidthPoint() > longestWordChunk.GetWidthPoint())
                    {
                        longestWord = word;
                    }
                }

                longestWordChunk = new Chunk(longestWord);
                longestWordChunk.Font = textChunk.Font;
                longestWordChunk.SetCharacterSpacing(textChunk.GetCharacterSpacing());

                return Stroke(longestWordChunk.GetWidthPoint(), -3f, 0, 0.5f, BaseColor.BLACK, alignment);
            }
            else
            {
                return Stroke(textChunk.GetWidthPoint(), -3f, 0, 0.5f, BaseColor.BLACK, alignment);
            }
        }

        static Font GetNormalFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.NORMAL, new BaseColor(0, 0, 0));
        }

        static Font GetBoldFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.BOLD, new BaseColor(0, 0, 0));
        }

        static Font GetItalicFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.ITALIC, new BaseColor(0, 0, 0));
        }

        static Font GetBoldItalicFont(float size)
        {
            return new Font(Font.FontFamily.HELVETICA, size, Font.BOLDITALIC, new BaseColor(0, 0, 0));
        }

        static string FormatAmount(decimal amount)
        {
            string formattedAmount = amount.ToString("N");

             //Puts negative sign at end
            if (formattedAmount.StartsWith("-"))
            {
              formattedAmount = formattedAmount.Replace("-", string.Empty);
              formattedAmount += "-";
            }

            return formattedAmount;
        }

        static DateTime ParseDate(string date)
        {
            DateTime parsedDate = new DateTime();

            date = date.Replace("-", string.Empty); // Removes dash from MoneyPerks file dates

            if (date.Length >= "MMDDYYYY".Length)
            {
                try
                {
                    int year = int.Parse(date.Substring("MMDD".Length, "YYYY".Length));
                    int month = int.Parse(date.Substring(0, "MM".Length));
                    int day = int.Parse(date.Substring("MM".Length, "DD".Length));
                    parsedDate = new DateTime(year, month, day);
                }
                catch (Exception exception)
                {
                    Log(exception.Message);
                }
            }

            return parsedDate;
        }

        private static void Log(string message)
        {
            try
            {
                StackTrace stackTrace = new StackTrace();
                StackFrame stackFrame = stackTrace.GetFrame(1);
                MethodBase methodBase = stackFrame.GetMethod();
                string methodName = methodBase.Name;
                int envelopeCount = MemberStatement.epilogue.accountCount;
                for (int i = 0; i < envelopeCount; i++)
                {
                    if (MemberStatement != null)
                    {
                        LogWriter.WriteLine(methodBase.DeclaringType.Name + "." + methodName + ": " + message + " (Account Number: " + MemberStatement.envelope[i].statement.account.accountNumber + ")");
                    }
                    else
                    {
                        LogWriter.WriteLine(methodBase.DeclaringType.Name + "." + methodName + ": " + message);
                    }
                }

            }
            catch (Exception)
            {
            }
        }


        private static int ParseMoneyPerksAmount(string amount)
        {
            int parsedAmount = 0;

            if (amount != string.Empty)
            {
                try
                {
                    parsedAmount = int.Parse(amount);
                }
                catch (Exception exception)
                {
                    Log(exception.Message + " " + amount);
                }
            }

            return parsedAmount;
        }

        private static List<string> CSVParser(string strInputString)
        {
            int intCounter = 0, intLenght;
            StringBuilder strElem = new StringBuilder();
            List<string> alParsedCsv = new List<string>();
            intLenght = strInputString.Length;
            strElem = strElem.Append("");
            int intCurrState = 0;
            int[][] aActionDecider = new int[9][];
            //Build the state array
            aActionDecider[0] = new int[4] { 2, 0, 1, 5 };
            aActionDecider[1] = new int[4] { 6, 0, 1, 5 };
            aActionDecider[2] = new int[4] { 4, 3, 3, 6 };
            aActionDecider[3] = new int[4] { 4, 3, 3, 6 };
            aActionDecider[4] = new int[4] { 2, 8, 6, 7 };
            aActionDecider[5] = new int[4] { 5, 5, 5, 5 };
            aActionDecider[6] = new int[4] { 6, 6, 6, 6 };
            aActionDecider[7] = new int[4] { 5, 5, 5, 5 };
            aActionDecider[8] = new int[4] { 0, 0, 0, 0 };
            for (intCounter = 0; intCounter < intLenght; intCounter++)
            {
                intCurrState = aActionDecider[intCurrState]
                                          [CSVParser_GetInputID(strInputString[intCounter])];
                //take the necessary action depending upon the state 
                CSVParser_PerformAction(ref intCurrState, strInputString[intCounter],
                             ref strElem, ref alParsedCsv);
            }
            //End of line reached, hence input ID is 3
            intCurrState = aActionDecider[intCurrState][3];
            CSVParser_PerformAction(ref intCurrState, '\0', ref strElem, ref alParsedCsv);
            return alParsedCsv;
        }

        private static int CSVParser_GetInputID(char chrInput)
        {
            if (chrInput == '"')
            {
                return 0;
            }
            else if (chrInput == ',')
            {
                return 1;
            }
            else
            {
                return 2;
            }
        }
        private static void CSVParser_PerformAction(ref int intCurrState, char chrInputChar,
                            ref StringBuilder strElem, ref List<string> alParsedCsv)
        {
            string strTemp = null;
            switch (intCurrState)
            {
                case 0:
                    //Separate out value to array list
                    strTemp = strElem.ToString();
                    alParsedCsv.Add(strTemp);
                    strElem = new StringBuilder();
                    break;
                case 1:
                case 3:
                case 4:
                    //accumulate the character
                    strElem.Append(chrInputChar);
                    break;
                case 5:
                    //End of line reached. Separate out value to array list
                    strTemp = strElem.ToString();
                    alParsedCsv.Add(strTemp);
                    break;
                case 6:
                    //Erroneous input. Reject line.
                    alParsedCsv.Clear();
                    break;
                case 7:
                    //wipe ending " and Separate out value to array list
                    strElem.Remove(strElem.Length - 1, 1);
                    strTemp = strElem.ToString();
                    alParsedCsv.Add(strTemp);
                    strElem = new StringBuilder();
                    intCurrState = 5;
                    break;
                case 8:
                    //wipe ending " and Separate out value to array list
                    strElem.Remove(strElem.Length - 1, 1);
                    strTemp = strElem.ToString();
                    alParsedCsv.Add(strTemp);
                    strElem = new StringBuilder();
                    //goto state 0
                    intCurrState = 0;
                    break;
            }
        }
        public static int GetNumberofPayments()
        {
            return NumberOfPayments;
        }

        public static int GetNumberOfStatementsBuilt()
        {
            return NumberOfStatementsBuilt;
        }

        public static int GetNumberOfWithdrawalsBuilt()
        {
            return NumberOfWithdrawals;
        }

        public static int GetNumberOfDeposits()
        {
            return NumberOfDeposits;
        }

       public static int GetNumberOfChecks()
        {
            return NumberOfChecks;
        }

        public static int GetNumberOfAdvances()
       {
           return NumberOfAdvances;
       }

        static Document Doc
        {
            get;
            set;
        }

        static Advertisement[] AdvertisementTop
        {
            get;
            set;
        }

        static Advertisement AdvertisementBottom
        {
            get;
            set;
        }

        static PdfWriter Writer
        {
            get;
            set;
        }

        private static int NumberOfAdvances
        {
            get;
            set;
        }

        private static int NumberOfPayments
        {
            get;
            set;
        }

        private static int NumberOfChecks
        {
            get;
            set;
        }

        private static int NumberOfStatementsBuilt
        {
            get;
            set;
        }

        private static int NumberOfDeposits
        {
            get;
            set;
        }

        private static int NumberOfWithdrawals
        {
            get;
            set;
        }

        static statementProduction MemberStatement
        {
            get;
            set;
        }

        static StreamWriter LogWriter
        {
            get;
            set;
        }


        

        static Dictionary<string, MoneyPerksStatement> MoneyPerksStatements
        {
            get;
            set;
        }

      

        public static string TEMP_FILE_NAME = "statement_pdf.temp";
        public const int MONEYPERKS_TRANSACTION_RECORD_FIELD_COUNT = 5;
        public const int MAX_RELATIONSHIP_BASED_LEVELS = 10;
    }
    
    class StatementPageEvent : PdfPageEventHelper
    {
        public override void OnStartPage(PdfWriter writer, Document Document)
        {
            string nextPageTemplate = Configuration.GetStatementTemplateFilePath();

            if (Document.PageNumber > 1)
            {
                Document.SetMargins(STATEMENT_MARGIN_SIDES, STATEMENT_MARGIN_SIDES, STATEMENT_MARGIN_TOP, STATEMENT_MARGIN_BOTTOM);
                Document.NewPage();

                using (FileStream templateInputStream = File.Open(nextPageTemplate, FileMode.Open))
                {
                    // Loads existing PDF
                    PdfReader reader = new PdfReader(templateInputStream);
                    PdfContentByte contentByte = writer.DirectContent;
                    PdfImportedPage page = writer.GetImportedPage(reader, 1);

                    // Copies first page of existing PDF into output PDF
                    Document.NewPage();
                    contentByte.AddTemplate(page, 0, 0);
                }
            }
            else
            {
                Document.SetMargins(STATEMENT_MARGIN_SIDES, STATEMENT_MARGIN_SIDES, FIRST_PAGE_STATEMENT_MARGIN_TOP, STATEMENT_MARGIN_BOTTOM);
            }
        }


        public static float FIRST_PAGE_STATEMENT_MARGIN_TOP = 12f;
        public static float STATEMENT_MARGIN_TOP = 70f;
        public static float STATEMENT_MARGIN_BOTTOM = 30f;
        public static float STATEMENT_MARGIN_SIDES = 12f;
    }
    
    
    }

