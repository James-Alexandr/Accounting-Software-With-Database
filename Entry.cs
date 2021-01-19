using System;
using System.Collections.Generic;
using System.Text;

namespace LeapLog
{
    public class Entry
    {
        //class properties
        public Entry() {
            Date = DateTime.Now;
        }
        public DateTime Date { get; private set; }

        public string Account1 { get; set; }

        public string Account2 { get; set; }

        public string Type1 { get; set; }

        public string Type2 { get; set; }

        public double Debit { get; set; }

        public double Credit { get; set; }

    }
}
