using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestingExecl.Model
{
    public class Account
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public long CreatedBy { get; set; }
        public long ModifiedBy { get; set; }

        public string EventName { get; set; }
        public string Location { get; set; }
        public double Load { get; set; }
        public double BasicSalary { get; set; }
        public double HRA { get; set; }
        public double TotalSalary { get; set; } //BasicSalary + HRA
        public double Expense { get; set; }

        public double BalanceAmount { get; set; }//TotalSalary - Expense

        public string Comments { get; set; }


        public IEnumerable<Account> GetData()
        {
            // Create a list of accounts.
            var TestAccounts = new List<Account>
              {
                new Account {Code ="001",Name="John",Description="Testing",CreatedBy=01012022,ModifiedBy=12012022,EventName="Report",Location="Uae",Load=234.456,BasicSalary=11000,HRA=2000,Expense=1000,Comments="please update your KYC" },
                new Account {Code ="002",Name="Derick",Description="Dev",CreatedBy=2021,ModifiedBy=2022,EventName="Report",Location="India",Load=78.456,BasicSalary=15000,HRA=2500,Expense=30000,Comments="Update Your Pan Details with Bank Account"},
                new Account {Code ="003",Name="Ashiq",Description="QA",CreatedBy=2021,ModifiedBy=2022,EventName="Transaction",Location="UK",Load=567.6,BasicSalary=12000,HRA=2500,Expense=1500,Comments="Account Disabled"},
                new Account {Code ="004",Name="Sanil",Description="Senior QA",CreatedBy=2021,ModifiedBy=2022,EventName="Transaction",Location="Austrsila",Load=5167.6,BasicSalary=22000,HRA=500,Expense=1500,Comments="NPA"},
                new Account {Code ="005",Name="Rinshad",Description="Dev",CreatedBy=2021,ModifiedBy=2022,EventName="Transaction",Location="UAE",Load=234.00,BasicSalary=16000,HRA=2800,Expense=1800,Comments="Password Reset"},
                new Account {Code ="006",Name="Amal",Description="HR",CreatedBy=2021,ModifiedBy=2022,EventName="Transaction",Location="USA",Load=976.8888,BasicSalary=40000,HRA=2500,Expense=1900,Comments="Account Disabled"},
                new Account {Code ="007",Name="Preena",Description="Accounting",CreatedBy=2021,ModifiedBy=2022,EventName="Transaction",Location="Oman",Load=765.111,BasicSalary=25000,HRA=3000,Expense=29000,Comments="Update KYC"},
                new Account {Code ="008",Name="Appu",Description="Trainee",CreatedBy=2021,ModifiedBy=2022,EventName="Transaction",Location="India",Load=347.22,BasicSalary=19000,HRA=250,Expense=3000,Comments="no comments"},
                new Account {Code ="009",Name="Ravi",Description="QA",CreatedBy=2021,ModifiedBy=2022,EventName="Transaction",Location="Kuwait",Load=1.90,BasicSalary=17000,HRA=540,Expense=2000,Comments="Disabled"},
                new Account {Code ="010",Name="Lahiz",Description="Traniee",CreatedBy=2021,ModifiedBy=2022,EventName="Transaction",Location="India",Load=10.89,BasicSalary=10000,HRA=6500,Expense=18000,Comments="Pan Requested"}
        };
            return TestAccounts;
        }
    }
}
