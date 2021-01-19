using System;
using System.Collections.Generic;
using System.Text;

namespace LeapLog
{
    public class Entry_tacc
    {
        //class properties
        public Entry_tacc()
        {
            Date = DateTime.Now;
            Debit = new List<double>();
            Credit = new List<double>();
        }
        public DateTime Date { get; set; }

        public string Account { get; set; }

        public string Type { get; set; }

        public List<double> Debit { get; private set; }

        public List<double> Credit { get; private set; }

        public double Balance { get; set; }

        public double TotalDebit { get; set; }

        public double TotalCredit { get; set; }

        public Entry_tacc Clone()
        {
            return new Entry_tacc()
            {
                Account = this.Account,
                Balance = this.Balance,
                Credit = this.Credit,
                Date = this.Date,
                Debit = this.Debit,
                Type = this.Type,
                TotalDebit = this.TotalDebit,
                TotalCredit = this.TotalCredit
            };
        }
     }
}
