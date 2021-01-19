using System;
using System.Collections.Generic;
using System.Text;

namespace LeapLog
{
    public class balance_data
    {
        public balance_data()
        {
            assetsList = new List<Entry_tacc>();
            loeList = new List<Entry_tacc>();
        }
        public double total_assets { get; set; }

        public double total_loe { get; set; }

        public List<Entry_tacc> assetsList { get; set; }

        public List<Entry_tacc> loeList { get; set; }
    }
}
