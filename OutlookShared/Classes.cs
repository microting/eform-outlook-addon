using eFormShared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookShared
{
    public class Appointment
    {
        #region var/pop
        public string GlobalId { get; set; }
        public DateTime Start { get; set; }
        public int Duration { get; set; }
        public string Subject { get; set; }
        public string Location { get; set; }
        public string Body { get; set; }

        public int TemplateId { get; set; }
        public List<int> SiteIds { get; set; }
        public bool Connected { get; set; }
        public string Title { get; set; }
        public string Info { get; set; }
        public List<string> Replacements { get; set; }
        public DateTime? ExpireAt { get; set; }
        public bool ColorRule { get; set; }
        public string MicrotingUId { get; set; }
   
        Tools t = new Tools();
        #endregion

        #region con
        public Appointment()
        {

        }

        public Appointment(string globalId, DateTime start, int duration, string subject, string location, string body, bool intrepidBody)
        {
            GlobalId = globalId;
            Start = start;
            Duration = duration;
            Subject = subject;
            Location = location;
            Body = body;

            TemplateId = -1;
            SiteIds = new List<int>();
            Connected = false;
            Title = "";
            Info = "";
            Replacements = new List<string>();
            ExpireAt = DateTime.Now.AddDays(2);
            ColorRule = true;
            MicrotingUId = "";

            if (intrepidBody)
                BodyToFields(body);
        }
        #endregion

        #region Public
        public void     BodyToFields(string body)
        {
            if (body == null)
                body = "";

            //KeyPoint
            try
            {
                string intrepidFailedStr = ReadingFields(body);

                if (intrepidFailedStr != "")
                {
                    Location = "Failed_to_intrepid";
                    Body =
                    "<<Info field: Intrepid error: Start>>" + Environment.NewLine +
                    intrepidFailedStr + Environment.NewLine +
                    "<<Info field: Intrepid error: End>>" + Environment.NewLine +
                    Environment.NewLine +
                    Body;
                }
            }
            catch (Exception ex)
            {
                Location = "Exception";
                Body =
                "<<Info field: Exception: Start>>" + Environment.NewLine +
                t.PrintException("Failed to intrepid this event, for the following reason:", ex) + Environment.NewLine +
                "<<Info field: Exception: End>>" + Environment.NewLine +
                Environment.NewLine +
                Body;
            }
        }

        private string   ReadingFields(string body)
        {
            string[] lines = body.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            SiteIds = new List<int>();
            string check = "";
            string rtrnMsg = "";

            foreach (var line in lines)
            {

                try
                {
                    string input = line.ToLower();

                    if (input == "")
                        continue;

                    check = "template:";
                    if (input.Contains(check))
                        TemplateId = int.Parse(input.Replace(check, "").Trim());

                    check = "sites:";
                    if (input.Contains(check))
                    {
                        string temp = input.Replace(check, "").Trim();

                        foreach (var item in t.TextLst(temp))
                        {
                            SiteIds.Add(int.Parse(item));
                        }
                    }
     
                    check = "connected:";
                    if (input.Contains(check))
                        Connected = t.Bool(input.Replace(check, "").Trim());
            
                    check = "title:";
                    if (input.Contains(check))
                        Title = input.Replace(check, "").Trim();
      
                    check = "info:";
                    if (input.Contains(check))
                        Info = input.Replace(check, "").Trim();

                    check = "expire:";
                    if (input.Contains(check))
                        ExpireAt = t.Date(input.Replace(check, "").Trim());

                    check = "color:";
                    if (input.Contains(check))
                        ColorRule = t.Bool(input.Replace(check, "").Trim());

                    check = "colour:";
                    if (input.Contains(check))
                        ColorRule = t.Bool(input.Replace(check, "").Trim());
                }
                catch { }
            }

            try
            { 
                Replacements = null;
            }
            catch { }


            //none-optional
            if (SiteIds.Count < 1)
                rtrnMsg = "The mandatory 'sites' input not recognized." + Environment.NewLine +
                    "- Expected format: sites:[identifier](,[identifier])*n" + Environment.NewLine +
                    "- Sample 1       : sites:1234,2345,3456" + Environment.NewLine +
                    "- Sample 2       : sites:'Salg',1234,'Peter',2345" + Environment.NewLine +
                    "" + Environment.NewLine + rtrnMsg;

            if (TemplateId < 1)
                rtrnMsg = "The mandatory 'template' input not recognized." + Environment.NewLine +
                    "- Expected format: template:[identifier]" + Environment.NewLine +
                    "- Sample 1       : template:12" + Environment.NewLine +
                    "- Sample 2       : template:'Container check'" + Environment.NewLine + 
                    "" + Environment.NewLine + rtrnMsg;

            return rtrnMsg.Trim();
        }

        public override string ToString()
        {
            string globalId = "";
            string title = "";

            if (GlobalId != null)
                globalId = GlobalId;

            if (Title != null)
                title = Title;

            return "GlobalId:" + globalId + " / Start:" + Start + " / Title:" + title;
        }
        #endregion
    }

    public enum WorkflowState
    {
        Planned,
        Processed,
        Created,
        Sent,
        Retrived,
        Completed,
        Canceled,
        Revoked,
        Failed_to_expection,
        Failed_to_intrepid
    }
}