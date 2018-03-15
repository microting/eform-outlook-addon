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
        public DateTime End { get; set; }
        public int Duration { get; set; }
        public int? Id { get; set; }
        public string Subject { get; set; }
        public string ProcessingState { get; set; }
        public string Body { get; set; }
        public bool Completed { get; set; }

        public int TemplateId { get; set; }
        public List<AppoinntmentSite> AppointmentSites { get; set; }
        public bool Connected { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
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

        public Appointment(string globalId, DateTime start, int duration, string subject, string processingState, string body, bool colorRule, bool parseBodyContent, int? id)
        {
            Id = id;
            GlobalId = globalId;
            Start = start;
            Duration = duration;
            End = start.AddMinutes(duration);
            Subject = subject;
            ProcessingState = processingState;
            Body = body;

            TemplateId = -1;
            AppointmentSites = new List<AppoinntmentSite>();
            Connected = false;
            Title = "";
            Description = "";
            Info = "";
            Replacements = new List<string>();
            Expire = 2;
            ColorRule = colorRule;
            MicrotingUId = "";

            if (parseBodyContent)
                BodyToFields(body);
        }
        #endregion

        #region public
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

            if (ProcessingState != null)
                location = ProcessingState;

            return "GlobalId:" + globalId + " / Start:" + start + " / Title:" + title + " / Location:" + location;
        }
        #endregion

        #region private
        private void BodyToFields(string body)
        {
            if (body == null)
                body = "";

            //KeyPoint
            try
            {
                string parsedFailedStr = ReadingFields(body);

                if (parsedFailedStr != "")
                {
                    ProcessingState = ProcessingStateOptions.ParsingFailed.ToString();
                    Body =
                    "<<< Interpret error: Start >>>" + Environment.NewLine +
                    parsedFailedStr + Environment.NewLine +
                    "<<< Interpret error: End >>>" + Environment.NewLine +
                    Environment.NewLine +
                    Body;
                }
            }
            catch (Exception ex)
            {
                ProcessingState = ProcessingStateOptions.Exception.ToString();
                Body =
                "<<< Exception: Start >>>" + Environment.NewLine +
                t.PrintException("Failed to intrepid this event, for the following reason:", ex) + Environment.NewLine +
                "<<< Exception: End >>>" + Environment.NewLine +
                Environment.NewLine +
                Body;
            }
        }

        private string ReadingFields(string body)
        {
            #region var
            string[] lines = body.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            AppointmentSites = new List<AppoinntmentSite>();
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
                            AppoinntmentSite appointmentSite = new AppoinntmentSite(null, int.Parse(item), ProcessingStateOptions.Processed.ToString(), null);
                            AppointmentSites.Add(appointmentSite);
                        }

                        AppointmentSites = AppointmentSites.Distinct().ToList();

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

                    check = "description#";
                    if (input.Contains(check))
                    {
                        string temp = line.Remove(0, check.Length).Trim();

                        if (Description == "")
                            Description = temp;
                        else
                            Description += "<br>" + temp;

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
            if (AppointmentSites.Count < 1)
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
        #endregion
    }

    public enum ProcessingStateOptions
    {
        //Appointment locations options / ProcessingState options
        Pre_created,
        Planned,
        Processed,
        Created,
        Sent,
        Retrived,
        Completed,
        Canceled,
        Revoked,
        ParsingFailed,
        Exception,
        Unknown_location
    }
}