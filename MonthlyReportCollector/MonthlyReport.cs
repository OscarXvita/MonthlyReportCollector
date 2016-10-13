using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MonthlyReportCollector
{
    class MonthlyReport
    {
            string iD;
            public string ID
            {
                get { return iD; }
                set
                {
                    if (!Regex.IsMatch(ID, @"[0-9]{4}"))
                    {
                        throw new ArgumentException("ID无效或为空！");

                    }
                    else
                    {
                        ID = iD;
                    }
                }
            }

            string Name;
            string Sex;
            string Team;
            string Phone;
            string Email;
            string City;
            string School;
            string Major;
            string Grade;
            int BlogNum;
            string BlogLink;
            int SocialNum;
            string SocialLink;
            int Retweets;
            string RTLink;
            int PostAccepted;
            string PostLink;
            int WindowsApps;
            string WALink;
            int ActivityHeldNum;
            string AHLink;
            int ActivityJoinNum;
            string AJNum;

        
    }
}
