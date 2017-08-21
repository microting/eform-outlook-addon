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
            Core core = new Core();

            string serverConnectionString = File.ReadAllText("input\\sql_connection.txt").Trim();

            while (true)
            {
                Console.WriteLine("");
                Console.WriteLine("Input:");
                Console.WriteLine("'S'tart, 'C'lose, 'R'eset & 'Q'uit");
                string tempLower = Console.ReadLine().ToLower();

                if (tempLower == "s")
                    core.Start(serverConnectionString);

                if (tempLower == "c")
                    core.Close();

                if (tempLower == "r")
                {
                    core.Close();
                    core.Test_Reset(serverConnectionString);
                }

                if (tempLower == "q")
                {
                    core.Close();
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