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

            string outConStr = "Data Source=.\\SQLEXPRESS;Initial Catalog=" + "MicrotingOutlook123" + ";Integrated Security=True";
            string sdkConStr = "Data Source=.\\SQLEXPRESS;Initial Catalog=" + "V166" + ";Integrated Security=True";
            //string serviceLocation = "";

            while (true)
            {
                #region text + read input
                Console.WriteLine("");
                Console.WriteLine("Input    : 'C'lose,'R'eset & 'Q'uit.");
                Console.WriteLine("Outlook  : 'O' start.  Running:" + outCore.Running());
                Console.WriteLine("SDK Core : 'S' start.  Running:" + sdkCore.Running());
                Console.WriteLine("-        : 'T'emplate");
                string tempLower = Console.ReadLine().ToLower();
                #endregion

                if (tempLower == "o")
                #region outlook core start
                {
                    outCore.Start(outConStr, GetServiceLocation());
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
                    //outCore.UnitTest_Reset(outConStr);
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

                if (tempLower == "t")
                #region template
                {
                    if (sdkCore.Running())
                    {
                        Console.WriteLine("Creating eForm template from the xmlTemplate.txt");

                        string xmlStr = File.ReadAllText("xmlTemplate.txt");
                        var main = sdkCore.TemplateFromXml(xmlStr);
                        main = sdkCore.TemplateUploadData(main);

                        // Best practice is to validate the parsed xml before trying to save and handle the error(s) gracefully.
                        List<string> validationErrors = sdkCore.TemplateValidation(main);
                        if (validationErrors.Count < 1)
                        {
                            main.Repeated = 1;
                            main.CaseType = "Test";
                            main.StartDate = DateTime.Now;
                            main.EndDate = DateTime.Now.AddDays(2);

                            try
                            {
                                Console.WriteLine("- TemplateId = " + sdkCore.TemplateCreate(main));
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("PANIC !!!", ex);
                            }
                        }
                        else
                        {
                            foreach (string error in validationErrors)
                                Console.WriteLine("The following error is stopping us from creating the template: " + error);

                            Console.WriteLine("Correct the errors in xmlTemplate.txt and try again");
                        }
                    }
                    else
                        Console.WriteLine("SDK Core not running");
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

        static string GetServiceLocation()
        {
            string serviceLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            serviceLocation = Path.GetDirectoryName(serviceLocation) + "\\";

            return serviceLocation;
        }
    }
}