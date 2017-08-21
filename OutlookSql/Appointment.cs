using eFormShared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookSql
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
        public int Expire { get; set; }
        public bool ColorRule { get; set; }
        public string MicrotingUId { get; set; }
   
        Tools t = new Tools();
        #endregion

        #region con
        public Appointment()
        {

        }

        public Appointment(string globalId, DateTime start, int duration, string subject, string location, string body, bool colorRule, bool intrepidBody, Func<string, string> Lookup)
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
            Expire = 2;
            ColorRule = colorRule;
            MicrotingUId = "";

            if (intrepidBody)
                BodyToFields(body, Lookup);
        }
        #endregion

        #region methods
        public void     BodyToFields(string body, Func<string, string> Lookup)
        {
            if (body == null)
                body = "";

            //KeyPoint
            try
            {
                string intrepidFailedStr = ReadingFields(body, Lookup);

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

        private string  ReadingFields(string body, Func<string, string> Lookup)
        {
            #region var
            string[] lines = body.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            SiteIds = new List<int>();
            Replacements = new List<string>();
            string check = "";
            string rtrnMsg = "";
            #endregion

            foreach (var line in lines)
            {
                try
                {
                    string input = line.ToLower();

                    #region template and sites
                    if (input.Trim() == "")
                        continue;

                    check = "template#";
                    if (input.Contains(check))
                    {
                        string itemStr = line.Remove(0, check.Length).Trim();

                        if (itemStr.Contains("'") || itemStr.Contains("’"))
                            itemStr = Lookup(itemStr.Replace("'", "").Replace("’", "").Trim());

                        if (itemStr.Contains("failed, for title"))
                            rtrnMsg = itemStr + Environment.NewLine +
                                "" + Environment.NewLine + rtrnMsg;
           
                        TemplateId = int.Parse(itemStr);

                        continue;
                    }

                    check = "sites#";
                    if (input.Contains(check))
                    {
                        string lineNoComma = line.Remove(0, check.Length).Trim();
                        lineNoComma = lineNoComma.Replace(",", "|");

                        foreach (var item in t.TextLst(lineNoComma))
                        {
                            if (item.Contains("'") || item.Contains("’"))
                            {
                                string itemStr = Lookup(item.Replace("'", "").Replace("’", "").Trim());

                                if (itemStr.Contains("failed, for title"))
                                    rtrnMsg = itemStr + Environment.NewLine +
                                        "" + Environment.NewLine + rtrnMsg;
                                else
                                    SiteIds.AddRange(t.IntLst(itemStr));
                            }
                            else
                                SiteIds.Add(int.Parse(item));
                        }

                        SiteIds = SiteIds.Distinct().ToList();

                        continue;
                    }
                    #endregion

                    #region tags
                    check = "connected#";
                    if (input.Contains(check))
                    {
                        Connected = t.Bool(line.Remove(0, check.Length).Trim());
                        continue;
                    }

                    check = "title#";
                    if (input.Contains(check))
                    {
                        Title = line.Remove(0, check.Length).Trim();
                        continue;
                    }

                    check = "info#";
                    if (input.Contains(check))
                    {
                        string temp = line.Remove(0, check.Length).Trim();

                        if (Info == "")
                            Info = temp;
                        else
                            Info += "<br>" + temp;

                        continue;
                    }

                    check = "expire#";
                    if (input.Contains(check))
                    {
                        Expire = int.Parse(line.Remove(0, check.Length).Trim());
                        continue;
                    }

                    check = "color#";
                    if (input.Contains(check))
                    {
                        ColorRule = t.Bool(line.Remove(0, check.Length).Trim());
                        continue;
                    }

                    check = "colour#";
                    if (input.Contains(check))
                    {
                        ColorRule = t.Bool(line.Remove(0, check.Length).Trim());
                        continue;
                    }

                    check = "replacements#";
                    if (input.Contains(check))
                    {
                        if (input.Contains("=="))
                        {
                            Replacements.Add(line.Remove(0, check.Length).Trim());
                            continue;
                        }
                        else
                        {
                            rtrnMsg = "The following replacement line:'" + line + "' did not contain a '=='." + Environment.NewLine +
                                "- Expected format: replacements#[old text]==[new text]" + Environment.NewLine +
                                "- Sample 1       : replacements#Location==Odense" + Environment.NewLine +
                                "- Sample 2       : replacements#[Choice1]==true" + Environment.NewLine +
                                "" + Environment.NewLine + rtrnMsg;
                            continue;
                        }
                    }
                    #endregion

                    //unknown
                    if (input.Contains("#"))
                        rtrnMsg = "The following line:'" + line + "' contains a '#'. Line tag not recognized." + Environment.NewLine +
                            "- Expected format: [line tag]#[infomation]" + Environment.NewLine +
                            "- Known line tags: 'connected', 'title', 'info', 'expire', 'color', 'colour' & 'replacements'" + Environment.NewLine +
                            "- Sample 1       : connected#false" + Environment.NewLine +
                            "- Sample 2       : connected# 0" + Environment.NewLine +
                            "" + Environment.NewLine + rtrnMsg;
                }
                catch (Exception ex)
                {
                        rtrnMsg = t.PrintException("The following line:'" + line + "' coursed a exception", ex) + Environment.NewLine +
                          "" + Environment.NewLine + rtrnMsg;
                }
            }

            #region none-optional
            if (SiteIds.Count < 1)
                rtrnMsg = "The mandatory field 'sites' input not recognized." + Environment.NewLine +
                    "- Expected format: sites#[identifier](,[identifier])*n" + Environment.NewLine +
                    "- Sample 1       : sites#1234,2345,3456" + Environment.NewLine +
                    "- Sample 2       : sites#'Salg',1234,'Peter',2345" + Environment.NewLine +
                    "" + Environment.NewLine + rtrnMsg;

            if (TemplateId < 1)
                rtrnMsg = "The mandatory field 'template' input not recognized." + Environment.NewLine +
                    "- Expected format: template#[identifier]" + Environment.NewLine +
                    "- Sample 1       : template#12" + Environment.NewLine +
                    "- Sample 2       : template#'Container check'" + Environment.NewLine + 
                    "" + Environment.NewLine + rtrnMsg;
            #endregion

            return rtrnMsg.Trim();
        }

        public override string ToString()
        {
            string globalId = "";
            string start = "";
            string title = "";
            string location = "";

            if (GlobalId != null)
                globalId = GlobalId;

            if (Start != null)
                start = Start.ToString();

            if (Title != null)
                title = Title;

            if (Location != null)
                location = Location;

            return "GlobalId:" + globalId + " / Start:" + start + " / Title:" + title + " / Location:" + location;
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