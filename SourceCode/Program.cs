using eFormCore;
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
            OutlookCore.Core oCore = new OutlookCore.Core();
            ICore iCore = new eFormCore.Core();

            string serverConnectionString = File.ReadAllText("input\\sql_connection.txt").Trim();
            string serverConnectionStrSdk = File.ReadAllText("input\\sql_connection_sdk.txt").Trim();
            oCore.HandleEventException += Stump;

            while (true)
            {
                Console.WriteLine("");
                Console.WriteLine("Input:");
                Console.WriteLine("'S'tart, 'C'lose, 'R'eset & 'Q'uit");
                string tempLower = Console.ReadLine().ToLower();

                if (tempLower == "s")
                {
                    oCore.Start(serverConnectionString);
                    iCore.Start(serverConnectionStrSdk);
                }

                if (tempLower == "c")
                {
                    oCore.Close();
                    iCore.Close();
                }

                if (tempLower == "r")
                {
                    oCore.Close();
                    iCore.Close();

                    oCore.Test_Reset(serverConnectionString);
                }

                if (tempLower == "q")
                {
                    oCore.Close();
                    iCore.Close();

                    break;
                }
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