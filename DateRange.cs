using System;
using System.Collections.Generic;
using System.Text;

namespace LeapLog
{
    public struct DateRange
    {
        public DateTime from_date;
        public DateTime to_date;

        public DateRange(DateTime from_date, DateTime to_date)
        {
            this.from_date = from_date;
            this.to_date = to_date;
        }
    }
}
