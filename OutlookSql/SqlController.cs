using eFormShared;
using OutlookShared;
using OutlookSql.Migrations;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.Migrations;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookSql
{
    public class SqlController
    {
        #region var
        string connectionStr;
        Tools t = new Tools();
        #endregion

        #region con
        public              SqlController(string connectionString)
        {
            try
            {
                if (string.IsNullOrEmpty(connectionString))
                    throw new ArgumentException("connectionString is not allowed to be null or empty");

                connectionStr = connectionString;
                PrimeDb(); //if needed
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        private void        PrimeDb()
        {
            int settingsCount = 0;

            try
            #region checks database connectionString works
            {
                using (var db = new OutlookDb(connectionStr))
                {
                    settingsCount = db.settings.Count();
                }
            }
            #endregion
            catch (Exception ex)
            #region if failed, will try to update context
            {
                if (ex.Message.Contains("context has changed") || ex.Message.Contains("'cases'"))
                {
                    var configuration = new Configuration();
                    configuration.TargetDatabase = new DbConnectionInfo(connectionStr, "System.Data.SqlClient");
                    var migrator = new DbMigrator(configuration);
                    migrator.Update();
                }
                else
                    throw;
            }
            #endregion

            if (SettingCheckAll())
                return;

            #region prime db
            try
            {
                using (var db = new OutlookDb(connectionStr))
                {
                    if (settingsCount != Enum.GetNames(typeof(Settings)).Length)
                    {
                        if (settingsCount == 0)
                            SettingPrime();
                        else
                            throw new Exception("FATAL Exception. Settings needs to be corrected. Please either inspect or clear the Settings table in the Microting database");
                    }
                }
            }
            catch (Exception ex)
            {
                // This is here because, the priming process of the DB, will require us to go through the process of migrating the DB multiple times.
                if (ex.Message.Contains("context has changed"))
                {
                    var configuration = new Configuration();
                    configuration.TargetDatabase = new DbConnectionInfo(connectionStr, "System.Data.SqlClient");
                    var migrator = new DbMigrator(configuration);
                    migrator.Update();
                    PrimeDb(); // It's on purpose we call our self until we have no more migrations.
                }
                else
                    throw new Exception(t.GetMethodName() + " failed", ex);
            }
            #endregion
        }
        #endregion

        #region public Outlook
        public bool         OutlookEfromCreate(Appointment appointment)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
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
                LogException(t.GetMethodName() + " failed", ex, false);
                return false;
            }
        }

        public bool         OutlookEformCancel(Appointment appointment)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
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
                LogException(t.GetMethodName() + " failed", ex, false);
                return false;
            }
        }

        public appointments AppointmentsFind(string globalId)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
                {
                    var match = db.appointments.SingleOrDefault(x => x.global_id == globalId);
                    return match;
                }
            }
            catch (Exception ex)
            {
                LogException(t.GetMethodName() + " failed", ex, false);
                return null;
            }
        }

        public appointments AppointmentsFindOne(WorkflowState workflowState)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
                {
                    var match = db.appointments.FirstOrDefault(x => x.workflow_state == workflowState.ToString());
                    return match;
                }
            }
            catch (Exception ex)
            {
                LogException(t.GetMethodName() + " failed", ex, false);
                return null;
            }
        }

        public appointments AppointmentsFindOne(int timesReflected)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
                {
                    var match = db.appointments.FirstOrDefault(x => x.completed == timesReflected);
                    return match;
                }
            }
            catch (Exception ex)
            {
                LogException(t.GetMethodName() + " failed", ex, false);
                return null;
            }
        }

        public bool         AppointmentsUpdate(string globalId, WorkflowState workflowState, string body, string expectionString, string response)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
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
                LogException(t.GetMethodName() + " failed", ex, false);
                return false;
            }
        }

        public bool         AppointmentsReflected(string globalId)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
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
                LogException(t.GetMethodName() + " failed", ex, false);
                return false;
            }
        }

        public string       Lookup(string title)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
                {
                    lookups match = db.lookups.Single(x => x.title == title);
                    return match.value;
                }
            }
            catch (Exception ex)
            {
                LogEverything(t.PrintException(t.GetMethodName() + " failed, for title:'" + title + "'", ex));
                return t.GetMethodName() + " failed, for title:'" + title + "'";
            }
        }
        #endregion

        #region public SDK interaction cases
        public bool         SyncInteractionCase()
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
                        LogVariable(nameof(appoint), appoint.ToString());
                        LogException("Failed to update Outlook appointment, but Appointment created in SDK input", new Exception("FATAL issue"), true);
                    }
                }
                else
                {
                    LogVariable(nameof(appoint), appoint.ToString());
                    LogException("Failed to created Appointment in SDK input", new Exception("FATAL issue"), true);
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
                        LogVariable(nameof(appoint), appoint.ToString());
                        LogException("Failed to update Outlook appointment, but Appointment deleted in SDK input", new Exception("FATAL issue"), true);
                    }
                }
                else
                {
                    LogVariable(nameof(appoint), appoint.ToString());
                    LogException("Failed to deleted Appointment in SDK input", new Exception("FATAL issue"), true);
                }

                return false;
            }
            #endregion

            // read output
            return InteractionCaseProcessed();
        }

        public bool         InteractionCaseCreate(appointments appointment)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
                {
                    List<int> siteIds = t.IntLst(appointment.site_ids);
                    List<string> replacements = t.TextLst(appointment.replacements);

                    if (replacements == null)
                        replacements = new List<string>();

                    if (appointment.title != "")
                        replacements.Add("Title::" + appointment.title);

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
                LogWarning(t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex));
                AppointmentsUpdate(appointment.global_id, WorkflowState.Failed_to_expection, appointment.body, t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex), null);
                return false;
            }
        }

        public bool         InteractionCaseDelete(appointments appointment)
        {
            try
            {
                eFormSqlController.SqlController sqlController = new eFormSqlController.SqlController(SettingRead(Settings.microtingDb));
                sqlController.InteractionCaseDelete(int.Parse(appointment.microting_uid));

                return true;
            }
            catch (Exception ex)
            {
                LogWarning(t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex));
                AppointmentsUpdate(appointment.global_id, WorkflowState.Failed_to_expection, appointment.body, t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex), null);
                return false;
            }
        }

        public bool         InteractionCaseProcessed()
        {
            try
            {
                using (var db = new MicrotingDb(SettingRead(Settings.microtingDb)))
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
                    int statEnd = 0;
                    string addToBody = "";
                    bool flagException = false;
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

                    if (match.workflow_state == "failed to sync")
                        flagException = true;

                    if (t.Bool(AppointmentsFind(match.custom).color_rule))
                        statEnd = statHigh;
                    else
                        statEnd = statLow;

                    #region WorkflowState wFS = ...
                    WorkflowState wFS = WorkflowState.Failed_to_intrepid;
                    if (statEnd == 1)
                        wFS = WorkflowState.Created;
                    if (statEnd == 2)
                        wFS = WorkflowState.Sent;
                    if (statEnd == 3)
                        wFS = WorkflowState.Retrived;
                    if (statEnd == 4)
                        wFS = WorkflowState.Completed;
                    if (statEnd == 5)
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
                LogException(t.GetMethodName() + " failed", ex, false);
                return true;
            }
        }
        #endregion

        #region public setting
        private void        SettingPrime()
        {
            using (var db = new OutlookDb(connectionStr))
            {
                SettingCreate(Settings.firstRunDone, 1);
                SettingCreate(Settings.logLevel, 2);
                SettingCreate(Settings.microtingDb, 3);
                SettingCreate(Settings.checkLast_At, 4);
                SettingCreate(Settings.checkRetrace_Hours, 5);
                SettingCreate(Settings.checkEvery_Mins, 6);
                SettingCreate(Settings.preSend_Mins, 7);
                SettingCreate(Settings.includeBlankLocations, 8);
                SettingCreate(Settings.colorsRule, 9);

                SettingUpdate(Settings.firstRunDone, "true");
                SettingUpdate(Settings.logLevel, "true");
                #region SettingUpdate(Settings.microtingDb, connectionStr.Replace("MicrotingOutlook", "Microting"));
                try
                {
                    SettingUpdate(Settings.microtingDb, connectionStr.Replace("MicrotingOutlook", "Microting"));
                }
                catch
                {
                    SettingUpdate(Settings.microtingDb, "xxxxx");
                }
                #endregion
                SettingUpdate(Settings.checkLast_At, DateTime.Now.AddMonths(-3).ToString());
                SettingUpdate(Settings.checkRetrace_Hours, "36");
                SettingUpdate(Settings.checkEvery_Mins, "15");
                SettingUpdate(Settings.preSend_Mins, "1");
                SettingUpdate(Settings.includeBlankLocations, "false");
                SettingUpdate(Settings.colorsRule, "1");
            }
        }

        public bool         SettingCheckAll()
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
                {
                    int countVal = db.settings.Count(x => x.value == "");
                    int countSet = db.settings.Count();

                    if (countVal > 0)
                        return false;

                    if (countSet < Enum.GetNames(typeof(Settings)).Length)
                        return false;

                    int failed = 0;
                    failed += SettingCheck(Settings.firstRunDone);
                    failed += SettingCheck(Settings.logLevel);
                    failed += SettingCheck(Settings.microtingDb);
                    failed += SettingCheck(Settings.checkLast_At);
                    failed += SettingCheck(Settings.checkRetrace_Hours);
                    failed += SettingCheck(Settings.checkEvery_Mins);
                    failed += SettingCheck(Settings.preSend_Mins);
                    failed += SettingCheck(Settings.includeBlankLocations);
                    failed += SettingCheck(Settings.colorsRule);
                    if (failed > 0)
                        return false;

                    return true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        private void        SettingCreate(Settings name, int id)
        {
            using (var db = new OutlookDb(connectionStr))
            {
                settings set = new settings();
                set.id = id;
                set.name = name.ToString();
                set.value = "";

                db.settings.Add(set);
                db.SaveChanges();
            }
        }

        public string       SettingRead(Settings name)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
                {
                    settings match = db.settings.SingleOrDefault(x => x.name == name.ToString());
                    return match.value;
                }
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public void         SettingUpdate(Settings name, string newValue)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
                {
                    settings match = db.settings.Single(x => x.name == name.ToString());
                    match.value = newValue;
                    db.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }
        #endregion

        #region public log
        public void     LogEverything(string str)
        {
            //Console.WriteLine("Lvl_(4): " + str);
        }

        public void     LogVariable(string variableName, string variableContent)
        {
            //Console.WriteLine("Lvl_(3): Variable Name:" + variableName + " / Content:" + variableContent);
        }

        public void     LogStandard(string str)
        {
            Console.WriteLine("Lvl_(2): " + str);
        }

        public void     LogCritical(string str)
        {
            Console.WriteLine("Lvl_(1): " + str);
        }

        public void     LogWarning(string str)
        {
            Console.WriteLine("Warn(0): " + str);
        }

        public void     LogException(string exceptionDescription, Exception exception, bool restartCore)
        {
            Console.WriteLine("Exp(-1): " + exceptionDescription);
            Console.WriteLine(t.PrintException(exceptionDescription, exception));

            //TODO - restart - do something
        }
        #endregion

        #region private
        private int         SettingCheck(Settings setting)
        {
            try
            {
                SettingRead(setting);
                return 0;
            }
            catch
            {
                return 1;
            }
        }

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
        public bool         UnitTest_TruncateTable_Outlook(string tableName)
        {
            try
            {
                using (var db = new OutlookDb(connectionStr))
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

        public bool         UnitTest_TruncateTable_Microting(string tableName)
        {
            try
            {
                using (var db = new MicrotingDb(SettingRead(Settings.microtingDb)))
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
        microtingDb,
        checkLast_At,
        checkRetrace_Hours,
        checkEvery_Mins,
        preSend_Mins,
        includeBlankLocations,
        colorsRule
    }
}