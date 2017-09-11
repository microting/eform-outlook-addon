using eFormData;
using OutlookCore;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SourceCode
{
    class Program
    {
        static void Main(string[] args)
        {
            Core outCore = new Core();
            eFormCore.Core sdkCore = new eFormCore.Core();

            string outConStr = "Data Source=DESKTOP-7V1APE5\\SQLEXPRESS;Initial Catalog=" + "MicrotingOutlook" + ";Integrated Security=True";
            string sdkConStr = "Data Source=DESKTOP-7V1APE5\\SQLEXPRESS;Initial Catalog=" + "MicrotingOdense"  + ";Integrated Security=True";
            
            while (true)
            {
                #region text + read input
                Console.WriteLine("");
                Console.WriteLine("Input    : 'C'lose,'R'eset & 'Q'uit.");
                Console.WriteLine("Outlook  : 'O' start.  Running:" + outCore.Running());
                Console.WriteLine("SDK Core : 'S' start.  Running:" + sdkCore.Running());
                string tempLower = Console.ReadLine().ToLower();
                #endregion

                if (tempLower == "o")
                #region outlook core start
                {
                    outCore.Start(outConStr);
                }
                #endregion

                if (tempLower == "s")
                #region SDK core start
                {
                    sdkCore.Start(sdkConStr);
                }
                #endregion

                if (tempLower == "r")
                #region reset
                {
                    sdkCore.Close();
                    outCore.Close();
                    outCore.Test_Reset(outConStr);
                }
                #endregion

                if (tempLower == "c")
                #region close
                {
                    sdkCore.Close();
                    outCore.Close();
                }
                #endregion

                if (tempLower == "q")
                #region quit
                {
                    sdkCore.Close();
                    outCore.Close();
                    break;
                }
                #endregion
            }

            Console.WriteLine("Closing app in 0,5 sec");
            Thread.Sleep(0500);
        }

        static void Stump(object sender, EventArgs args)
        {
            Console.WriteLine("EVENT:" + sender.ToString());
        }
    }
}