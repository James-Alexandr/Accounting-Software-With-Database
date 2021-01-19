using System;
using System.Collections.Generic;
using System.Text;

namespace LeapLog
{
    public class income_data
    {
        public income_data()
        {
            revenueList = new List<Entry_tacc>();
            expenseList = new List<Entry_tacc>();
        }
        public List<Entry_tacc> revenueList { get; set; }

        public List<Entry_tacc> expenseList { get; set; }

        public double total_revenue { get; set; }

        public double total_expenses { get; set; }

        public double net_income { get; set; }

        
    }
}
