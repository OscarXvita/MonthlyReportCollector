﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MonthlyReportCollector
{
    public class MonthlyReport
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Sex { get; set; }
        public string Team { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string City { get; set; }
        public string School { get; set; }
        public string Major { get; set; }
        public string Grade { get; set; }
        public int BlogNum { get; set; }
        public string BlogLink { get; set; }
        public int SocialNum { get; set; }
        public string SocialLink { get; set; }
        public int Retweets { get; set; }
        public string RtLink { get; set; }
        public int PostAccepted { get; set; }
        public string PostLink { get; set; }
        public int WindowsApps { get; set; }
        public string WaLink { get; set; }
        public int ActivityHeldNum { get; set; }
        public string AhLink { get; set; }
        public int ActivityJoinNum { get; set; }
        public string AjNum { get; set; }


    }
}
