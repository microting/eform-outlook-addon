using eFormShared;
using OutlookSql.Migrations;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.Migrations;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;

namespace OutlookSql
{
    public class SqlController : LogWriter
    {
        #region var
        string connectionStr;
        bool msSql = true;
        Log log;
        Tools t = new Tools();

        object _writeLock = new object();
        #endregion

        #region con
        public                      SqlController(string connectionString)
        {
            connectionStr = connectionString.ToLower();

            if (!connectionStr.Contains("server="))
                msSql = true;
            else
                msSql = false;

            #region migrate if needed
            try
            {
                using (var db = GetContextO())
                {
                    db.Database.CreateIfNotExists();
                    var match = db.settings.Count();
                }
            }
            catch
            {
                MigrateDb();
            }
            #endregion

            //region set default for settings if needed
            if (SettingCheckAll().Count > 0)
                SettingCreateDefaults();
        }

        private OutlookContextInterface   GetContextO()
        {
            if (msSql)
                return new OutlookDbMs(connectionStr);
            else
                return new OutlookDbMy(connectionStr);
        }

        private MicrotingContextInterface GetContextM()
        {
            if (msSql)
                return new MicrotingDbMs(SettingRead(Settings.microtingDb));
            else
                return new MicrotingDbMy(SettingRead(Settings.microtingDb));
        }

        public bool                 MigrateDb()
        {
            var configuration = new Configuration();
            configuration.TargetDatabase = new DbConnectionInfo(connectionStr, "System.Data.SqlClient");
            var migrator = new DbMigrator(configuration);
            migrator.Update();
            return true;
        }
        #endregion

        #region public
        #region public Outlook
        public bool                 OutlookEfromCreate(Appointment appointment)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.SingleOrDefault(x => x.global_id == appointment.GlobalId);

                    if (match != null)
                        return false;

                    appointments newAppo = new appointments();

                    newAppo.connected = t.Bool(appointment.Connected);
                    newAppo.replacements = t.TextLst(appointment.Replacements);
                    newAppo.duration = appointment.Duration;
                    newAppo.expire_at = appointment.Start.AddDays(appointment.Expire);
                    newAppo.global_id = appointment.GlobalId;
                    newAppo.info = appointment.Info;
                    newAppo.location = appointment.Location;
                    newAppo.body = appointment.Body;
                    newAppo.microting_uid = appointment.MicrotingUId;
                    newAppo.completed = 1;
                    newAppo.site_ids = t.IntLst(appointment.SiteIds);
                    newAppo.start_at = appointment.Start;
                    newAppo.subject = appointment.Subject;
                    newAppo.template_id = appointment.TemplateId;
                    newAppo.title = appointment.Title;
                    newAppo.description = appointment.Description;
                    newAppo.color_rule = t.Bool(appointment.ColorRule);
                    newAppo.workflow_state = "Processed";
                    newAppo.created_at = DateTime.Now;
                    newAppo.updated_at = DateTime.Now;
                    newAppo.version = 1;

                    db.appointments.Add(newAppo);
                    db.SaveChanges();

                    db.appointment_versions.Add(MapAppointmentVersions(newAppo));
                    db.SaveChanges();

                    return true;
                }
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, false);
                return false;
            }
        }

        public bool                 OutlookEformCancel(Appointment appointment)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.SingleOrDefault(x => x.global_id == appointment.GlobalId);

                    if (match == null)
                        return false;

                    match.workflow_state = "Canceled";
                    match.updated_at = DateTime.Now;
                    match.completed = 1;
                    match.version = match.version + 1;

                    db.SaveChanges();

                    db.appointment_versions.Add(MapAppointmentVersions(match));
                    db.SaveChanges();

                    return true;
                }
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, false);
                return false;
            }
        }

        public appointments         AppointmentsFind(string globalId)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.SingleOrDefault(x => x.global_id == globalId);
                    return match;
                }
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, false);
                return null;
            }
        }

        public appointments         AppointmentsFindOne(WorkflowState workflowState)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.FirstOrDefault(x => x.workflow_state == workflowState.ToString());
                    return match;
                }
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, false);
                return null;
            }
        }

        public appointments         AppointmentsFindOne(int timesReflected)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.FirstOrDefault(x => x.completed == timesReflected);
                    return match;
                }
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, false);
                return null;
            }
        }

        public bool                 AppointmentsUpdate(string globalId, WorkflowState workflowState, string body, string expectionString, string response)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.SingleOrDefault(x => x.global_id == globalId);

                    if (match == null)
                        return false;

                    match.workflow_state = workflowState.ToString();
                    match.updated_at = DateTime.Now;
                    match.completed = 0;
                    #region match.body = body ...
                    if (body != null)
                        match.body = body;
                    #endregion
                    match.response = response;
                    match.expectionString = expectionString;
                    match.version = match.version + 1;

                    db.SaveChanges();

                    db.appointment_versions.Add(MapAppointmentVersions(match));
                    db.SaveChanges();

                    return true;
                }
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, false);
                return false;
            }
        }

        public bool                 AppointmentsReflected(string globalId)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.SingleOrDefault(x => x.global_id == globalId);

                    if (match == null)
                        return false;

                    short temp = 0;

                    if (match.completed == 0)
                        temp = 1;

                    if (match.completed == 1)
                        temp = 2;

                    match.updated_at = DateTime.Now;
                    match.completed = temp;
                    match.version = match.version + 1;

                    db.SaveChanges();

                    db.appointment_versions.Add(MapAppointmentVersions(match));
                    db.SaveChanges();

                    return true;
                }
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, false);
                return false;
            }
        }

        public string               Lookup(string title)
        {
            try
            {
                using (var db = GetContextO())
                {
                    lookups match = db.lookups.Single(x => x.title == title);
                    return match.value;
                }
            }
            catch (Exception ex)
            {
                log.LogEverything("Not Specified", t.PrintException(t.GetMethodName() + " failed, for title:'" + title + "'", ex));
                return t.GetMethodName() + " failed, for title:'" + title + "'";
            }
        }
        #endregion

        #region public SDK interaction cases
        public bool                 SyncInteractionCase()
        {
            // read input
            #region create
            appointments appoint = AppointmentsFindOne(WorkflowState.Processed);

            if (appoint != null)
            {
                if (InteractionCaseCreate(appoint))
                {
                    bool isUpdated = AppointmentsUpdate(appoint.global_id, WorkflowState.Created, appoint.body, appoint.expectionString, null);

                    if (isUpdated)
                        return true;
                    else
                    {
                        log.LogVariable("Not Specified", nameof(appoint), appoint.ToString());
                        log.LogException("Not Specified", "Failed to update Outlook appointment, but Appointment created in SDK input", new Exception("FATAL issue"), true);
                    }
                }
                else
                {
                    log.LogVariable("Not Specified", nameof(appoint), appoint.ToString());
                    log.LogException("Not Specified", "Failed to created Appointment in SDK input", new Exception("FATAL issue"), true);
                }

                return false;
            }
            #endregion

            #region delete
            appoint = AppointmentsFindOne(WorkflowState.Canceled);

            if (appoint != null)
            {
                if (InteractionCaseDelete(appoint))
                {
                    bool isUpdated = AppointmentsUpdate(appoint.global_id, WorkflowState.Revoked, appoint.body, appoint.expectionString, null);

                    if (isUpdated)
                        return true;
                    else
                    {
                        log.LogVariable("Not Specified", nameof(appoint), appoint.ToString());
                        log.LogException("Not Specified", "Failed to update Outlook appointment, but Appointment deleted in SDK input", new Exception("FATAL issue"), true);
                    }
                }
                else
                {
                    log.LogVariable("Not Specified", nameof(appoint), appoint.ToString());
                    log.LogException("Not Specified", "Failed to deleted Appointment in SDK input", new Exception("FATAL issue"), true);
                }

                return false;
            }
            #endregion

            // read output
            return InteractionCaseProcessed();
        }

        public bool                 InteractionCaseCreate(appointments appointment)
        {
            try
            {
                using (var db = GetContextO())
                {
                    List<int> siteIds = t.IntLst(appointment.site_ids);
                    List<string> replacements = t.TextLst(appointment.replacements);

                    if (replacements == null)
                        replacements = new List<string>();

                    if (appointment.title != "")
                        replacements.Add("Title::" + appointment.title);

                    if (appointment.description != "")
                        replacements.Add("Description::" + appointment.description);

                    if (appointment.info != "")
                        replacements.Add("Info::" + appointment.info);

                    if (appointment.expire_at != DateTime.MinValue)
                        replacements.Add("Expire::" + appointment.expire_at.ToString());

                    if (replacements.Count == 0)
                        replacements = null;

                    eFormSqlController.SqlController sqlController = new eFormSqlController.SqlController(SettingRead(Settings.microtingDb));
                    int interCaseId = sqlController.InteractionCaseCreate((int)appointment.template_id, "", siteIds, appointment.global_id, t.Bool(appointment.connected), replacements);

                    var match = db.appointments.Single(x => x.global_id == appointment.global_id);

                    match.microting_uid = "" + interCaseId;
                    match.updated_at = DateTime.Now;
                    match.version = match.version + 1;

                    db.SaveChanges();

                    db.appointment_versions.Add(MapAppointmentVersions(match));
                    db.SaveChanges();
                }

                return true;
            }
            catch (Exception ex)
            {
                log.LogWarning("Not Specified", t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex));
                AppointmentsUpdate(appointment.global_id, WorkflowState.Failed_to_expection, appointment.body, t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex), null);
                return false;
            }
        }

        public bool                 InteractionCaseDelete(appointments appointment)
        {
            try
            {
                eFormSqlController.SqlController sqlController = new eFormSqlController.SqlController(SettingRead(Settings.microtingDb));
                sqlController.InteractionCaseDelete(int.Parse(appointment.microting_uid));

                return true;
            }
            catch (Exception ex)
            {
                log.LogWarning("Not Specified", t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex));
                AppointmentsUpdate(appointment.global_id, WorkflowState.Failed_to_expection, appointment.body, t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex), null);
                return false;
            }
        }

        public bool                 InteractionCaseProcessed()
        {
            try
            {
                using (var db = GetContextM())
                {
                    var match = db.a_interaction_cases.FirstOrDefault(x => x.synced == 0);
                    if (match == null)
                        return false;

                    match.updated_at = DateTime.Now;
                    match.version = match.version++;
                    match.synced = 1;
                    db.SaveChanges();

                    #region var
                    int statHigh = -99;
                    int statLow = 99;
                    int statCur = 0;
                    int statFinal = 0;
                    string addToBody = "";
                    bool flagException = false;
                    bool anyCompleted = false;
                    #endregion
                    foreach (var item in match.a_interaction_case_lists)
                    {
                        #region if stat ...
                        statCur = 0;

                        if (item.stat == "Created")
                        {
                            statCur = 1;
                            addToBody += item.siteId + "/" + item.updated_at + "/" + item.stat + "/" + Environment.NewLine;
                        }
                        if (item.stat == "Sent")
                        {
                            statCur = 2;
                            addToBody += item.siteId + "/" + item.updated_at + "/" + item.stat + "/" + item.microting_uid + Environment.NewLine;
                        }
                        if (item.stat == "Retrived")
                        {
                            statCur = 3;
                            addToBody += item.siteId + "/" + item.updated_at + "/" + item.stat + "/" + item.microting_uid + Environment.NewLine;
                        }
                        if (item.stat == "Completed")
                        {
                            statCur = 4;
                            addToBody += item.siteId + "/" + item.updated_at + "/" + item.stat + "/" + item.microting_uid + "/" + item.check_uid + Environment.NewLine;
                            anyCompleted = true;
                        }
                        if (item.stat == "Deleted")
                        {
                            statCur = 5;
                            addToBody += item.siteId + "/" + item.updated_at + "/" + item.stat + Environment.NewLine;
                        }
                        
                        if (item.stat == "Expection")
                        {
                            flagException = true;
                            addToBody += item.siteId + "/" + item.updated_at + "/Exception" + Environment.NewLine;
                        }

                        if (statHigh < statCur)
                            statHigh = statCur;

                        if (statLow > statCur)
                            statLow = statCur;
                        #endregion
                    }

                    if (anyCompleted && statHigh == 5) //as in 1 or more completed, and some deleted
                        statHigh = 4;

                    if (match.workflow_state == "failed to sync")
                        flagException = true;

                    if (t.Bool(AppointmentsFind(match.custom).color_rule))
                        statFinal = statHigh;
                    else
                        statFinal = statLow;

                    #region WorkflowState wFS = ...
                    WorkflowState wFS = WorkflowState.Failed_to_intrepid;
                    if (statFinal == 1)
                        wFS = WorkflowState.Created;
                    if (statFinal == 2)
                        wFS = WorkflowState.Sent;
                    if (statFinal == 3)
                        wFS = WorkflowState.Retrived;
                    if (statFinal == 4)
                        wFS = WorkflowState.Completed;
                    if (statFinal == 5)
                        wFS = WorkflowState.Revoked;
                    if (flagException == true)
                        wFS = WorkflowState.Failed_to_intrepid;
                    #endregion

                    if (addToBody != "")
                        AppointmentsUpdate(match.custom, wFS, null, match.expectionString, addToBody);
                    else
                        AppointmentsUpdate(match.custom, wFS, null, match.expectionString, null);

                    return true;
                }
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, false);
                return true;
            }
        }
        #endregion

        #region public setting
        public bool                 SettingCreateDefaults()
        {
            //key point
            SettingCreate(Settings.firstRunDone);
            SettingCreate(Settings.logLevel);
            SettingCreate(Settings.logLimit);
            SettingCreate(Settings.microtingDb);
            SettingCreate(Settings.checkLast_At);
            SettingCreate(Settings.checkPreSend_Hours);
            SettingCreate(Settings.checkRetrace_Hours);
            SettingCreate(Settings.checkEvery_Mins);
            SettingCreate(Settings.includeBlankLocations);
            SettingCreate(Settings.colorsRule);
            SettingCreate(Settings.calendarName);

            return true;
        }

        public bool                 SettingCreate(Settings name)
        {
            using (var db = GetContextO())
            {
                //key point
                #region id = settings.name
                int id = -1;
                string defaultValue = "default";
                switch (name)
                {
                    case Settings.firstRunDone:             id =  1;    defaultValue = "false";                                 break;
                    case Settings.logLevel:                 id =  2;    defaultValue = "4";                                     break;
                    case Settings.logLimit:                 id =  3;    defaultValue = "250";                                   break;
                    #region  case Settings.microtingDb:              id =  4;    defaultValue = 'MicrotingDB';                           break;
                    case Settings.microtingDb:

                        string microtingConnectionString = "...missing...";
                        try
                        {
                            microtingConnectionString = connectionStr.Replace("MicrotingOutlook", "Microting");
                            SettingUpdate(Settings.firstRunDone, "true");
                        }
                        catch { }
                                                            id =  4;    defaultValue = microtingConnectionString;               break;
                    #endregion
                    case Settings.checkLast_At:             id =  5;    defaultValue = DateTime.Now.AddMonths(-3).ToString();   break;
                    case Settings.checkPreSend_Hours:       id =  6;    defaultValue = "36";                                    break;
                    case Settings.checkRetrace_Hours:       id =  7;    defaultValue = "36";                                    break;
                    case Settings.checkEvery_Mins:          id =  8;    defaultValue = "15";                                    break;
                    case Settings.includeBlankLocations:    id =  9;    defaultValue = "true";                                  break;
                    case Settings.colorsRule:               id = 10;    defaultValue = "1";                                     break;
                    case Settings.calendarName:             id = 11;    defaultValue = "default";                               break;
       
                    default:
                        throw new IndexOutOfRangeException(name.ToString() + " is not a known/mapped Settings type");
                }
                #endregion

                settings matchId = db.settings.SingleOrDefault(x => x.id == id);
                settings matchName = db.settings.SingleOrDefault(x => x.name == name.ToString());

                if (matchName == null)
                {
                    if (matchId != null)
                    {
                        #region there is already a setting with that id but different name
                        //the old setting data is copied, and new is added
                        settings newSettingBasedOnOld = new settings();
                        newSettingBasedOnOld.id = (db.settings.Select(x => (int?)x.id).Max() ?? 0) + 1;
                        newSettingBasedOnOld.name = matchId.name.ToString();
                        newSettingBasedOnOld.value = matchId.value;

                        db.settings.Add(newSettingBasedOnOld);

                        matchId.name = name.ToString();
                        matchId.value = defaultValue;

                        db.SaveChanges();
                        #endregion
                    }
                    else
                    {
                        //its a new setting
                        settings newSetting = new settings();
                        newSetting.id = id;
                        newSetting.name = name.ToString();
                        newSetting.value = defaultValue;

                        db.settings.Add(newSetting);
                    }
                    db.SaveChanges();
                }
                else
                    if (string.IsNullOrEmpty(matchName.value))
                        matchName.value = defaultValue;
            }

            return true;
        }

        public string               SettingRead(Settings name)
        {
            try
            {
                using (var db = GetContextO())
                {
                    settings match = db.settings.Single(x => x.name == name.ToString());

                    if (match.value == null)
                        return "";

                    return match.value;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public void                 SettingUpdate(Settings name, string newValue)
        {
            try
            {
                using (var db = GetContextO())
                {
                    settings match = db.settings.SingleOrDefault(x => x.name == name.ToString());

                    if (match == null)
                    {
                        SettingCreate(name);
                        match = db.settings.Single(x => x.name == name.ToString());
                    }

                    match.value = newValue;
                    db.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public List<string>         SettingCheckAll()
        {
            List<string> result = new List<string>();
            try
            {
                using (var db = GetContextO())
                {
                    int countVal = db.settings.Count(x => x.value == "");
                    int countSet = db.settings.Count();

                    if (countSet == 0)
                    {
                        result.Add("NO SETTINGS PRESENT, NEEDS PRIMING!");
                        return result;
                    }

                    foreach (var setting in Enum.GetValues(typeof(Settings)))
                    {
                        try
                        {
                            string readSetting = SettingRead((Settings)setting);
                            if (readSetting == "")
                                result.Add(setting.ToString() + " has an empty value!");
                        }
                        catch
                        {
                            result.Add("There is no setting for " + setting + "! You need to add one");
                        }
                    }
                    return result;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }
        #endregion

        #region public write log
        public Log                  StartLog(CoreBase core)
        {
            try
            {
                string logLevel = SettingRead(Settings.logLevel);
                int logLevelInt = int.Parse(logLevel);
                log = new Log(core, this, logLevelInt);
                return log;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public override string      WriteLogEntry(LogEntry logEntry)
        {
            lock (_writeLock)
            {
                try
                {
                    using (var db = GetContextO())
                    {
                        logs newLog = new logs();
                        newLog.created_at = logEntry.Time;
                        newLog.level = logEntry.Level;
                        newLog.message = logEntry.Message;
                        newLog.type = logEntry.Type;

                        db.logs.Add(newLog);
                        db.SaveChanges();

                        if (logEntry.Level < 0)
                            WriteLogExceptionEntry(logEntry);

                        #region clean up of log table
                        int limit = t.Int(SettingRead(Settings.logLimit));
                        if (limit > 0)
                        {
                            List<logs> killList = db.logs.Where(x => x.id <= newLog.id - limit).ToList();

                            if (killList.Count > 0)
                            {
                                db.logs.RemoveRange(killList);
                                db.SaveChanges();
                            }
                        }
                        #endregion
                    }
                    return "";
                }
                catch (Exception ex)
                {
                    return t.PrintException(t.GetMethodName() + " failed", ex);
                }
            }
        }

        private string              WriteLogExceptionEntry(LogEntry logEntry)
        {
            try
            {
                using (var db = GetContextO())
                {
                    log_exceptions newLog = new log_exceptions();
                    newLog.created_at = logEntry.Time;
                    newLog.level = logEntry.Level;
                    newLog.message = logEntry.Message;
                    newLog.type = logEntry.Type;

                    db.log_exceptions.Add(newLog);
                    db.SaveChanges();

                    #region clean up of log exception table
                    int limit = t.Int(SettingRead(Settings.logLimit));
                    if (limit > 0)
                    {
                        List<log_exceptions> killList = db.log_exceptions.Where(x => x.id <= newLog.id - limit).ToList();

                        if (killList.Count > 0)
                        {
                            db.log_exceptions.RemoveRange(killList);
                            db.SaveChanges();
                        }
                    }
                    #endregion
                }
                return "";
            }
            catch (Exception ex)
            {
                return t.PrintException(t.GetMethodName() + " failed", ex);
            }
        }

        public override void        WriteIfFailed(string logEntries)
        {
            lock (_writeLock)
            {
                try
                {
                    File.AppendAllText(@"expection.txt",
                        DateTime.Now.ToString() + " // " + "L:" + "-22" + " // " + "Write logic failed" + " // " + Environment.NewLine
                        + logEntries + Environment.NewLine);
                }
                catch
                {
                    //magic
                }
            }
        }
        #endregion
        #endregion

        #region private
        private appointment_versions MapAppointmentVersions(appointments appointment)
        {
            appointment_versions version = new appointment_versions();

            version.workflow_state = appointment.workflow_state;
            version.version = appointment.version;
            version.created_at = appointment.created_at;
            version.updated_at = appointment.updated_at;
            version.global_id = appointment.global_id;
            version.start_at = appointment.start_at;
            version.expire_at = appointment.expire_at;
            version.duration = appointment.duration;
            version.template_id = appointment.template_id;
            version.subject = appointment.subject;
            version.location = appointment.location;
            version.body = appointment.body;
            version.expectionString = appointment.expectionString;
            version.site_ids = appointment.site_ids;
            version.title = appointment.title;
            version.description = appointment.description;
            version.info = appointment.info;
            version.replacements = appointment.replacements;
            version.microting_uid = appointment.microting_uid;
            version.connected = appointment.connected;
            version.completed = appointment.completed;
            version.response_text = appointment.response;
            version.color_rule = appointment.color_rule;

            version.appointment_id = appointment.id; //<<--

            return version;
        }
        #endregion

        #region unit test
        public bool                 UnitTest_TruncateTable_Outlook(string tableName)
        {
            try
            {
                using (var db = GetContextO())
                {
                    db.Database.ExecuteSqlCommand("DELETE FROM [dbo].[" + tableName + "];");
                    db.Database.ExecuteSqlCommand("DBCC CHECKIDENT('" + tableName + "', RESEED, 0);");

                    return true;
                }
            }
            catch (Exception ex)
            {
                string str = ex.Message;
                return false;
            }
        }

        public bool                 UnitTest_TruncateTable_Microting(string tableName)
        {
            try
            {
                using (var db = GetContextM())
                {
                    db.Database.ExecuteSqlCommand("DELETE FROM [dbo].[" + tableName + "];");
                    db.Database.ExecuteSqlCommand("DBCC CHECKIDENT('" + tableName + "', RESEED, 0);");

                    return true;
                }
            }
            catch (Exception ex)
            {
                string str = ex.Message;
                return false;
            }
        }
        #endregion
    }

    public enum Settings
    {
        firstRunDone,
        logLevel,
        logLimit,
        microtingDb,
        checkLast_At,
        checkPreSend_Hours,
        checkRetrace_Hours,
        checkEvery_Mins,
        includeBlankLocations,
        colorsRule,
        calendarName
    }
}