﻿using eFormShared;
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
        eFormSqlController.SqlController sdkSqlCon = null;
        Log log = null;
        Tools t = new Tools();

        object _writeLock = new object();

        string connectionStr;
        bool msSql = true;
        #endregion

        #region con
        public                      SqlController(string connectionStringOutlook)
        {
            ConstructorBase(connectionStringOutlook);
        }

        public                      SqlController(string connectionStringOutlook, string connectionStringSdk)
        {
            try
            {
                using (var db = GetContextO())
                {
                    db.Database.CreateIfNotExists();
                }
            }
            catch
            {
                throw new Exception("Failed to create Outlook database");
            }

            sdkSqlCon = new eFormSqlController.SqlController(SettingRead(Settings.microtingDb));
            ConstructorBase(connectionStringOutlook);
        }

        private void                ConstructorBase(string connectionString)
        {
            connectionStr = connectionString;

            if (connectionStr.ToLower().Contains("uid=") || connectionStr.ToLower().Contains("pwd="))
                msSql = false;
            else
                msSql = true;

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

            sdkSqlCon = new eFormSqlController.SqlController(SettingRead(Settings.microtingDb));
        }

        private OutlookContextInterface     GetContextO()
        {
            if (msSql)
                return new OutlookDbMs(connectionStr);
            else
                return new OutlookDbMy(connectionStr);
        }

        private MicrotingContextInterface   GetContextM()
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
        public int                  AppointmentsCreate(Appointment appointment)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.SingleOrDefault(x => x.global_id == appointment.GlobalId);

                    if (match != null)
                        return -1;

                    appointments newAppo = new appointments();

                    newAppo.connected = t.Bool(appointment.Connected);
                    newAppo.replacements = t.TextLst(appointment.Replacements);
                    newAppo.duration = appointment.Duration;
                    newAppo.expire_at = appointment.Start.AddDays(appointment.Expire);
                    newAppo.global_id = appointment.GlobalId;
                    newAppo.info = appointment.Info;
                    newAppo.location = "Processed";
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
                    newAppo.workflow_state = "Created";
                    newAppo.created_at = DateTime.Now;
                    newAppo.updated_at = DateTime.Now;
                    newAppo.version = 1;

                    db.appointments.Add(newAppo);
                    db.SaveChanges();

                    db.appointment_versions.Add(MapAppointmentVersions(newAppo));
                    db.SaveChanges();

                    return newAppo.id;
                }
            }
            catch (Exception ex)
            {
                log.LogException("Not Specified", t.GetMethodName() + " failed", ex, false);
                return 0;
            }
        }

        public bool                 AppointmentsCancel(string globalId)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.SingleOrDefault(x => x.global_id == globalId);

                    if (match == null)
                        return false;

                    match.location = "Canceled";
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

        public appointments         AppointmentsFindOne(LocationOptions location)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.FirstOrDefault(x => x.location == location.ToString());
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

        public bool                 AppointmentsUpdate(string globalId, LocationOptions location, string body, string expectionString, string response)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.SingleOrDefault(x => x.global_id == globalId);

                    if (match == null)
                        return false;

                    match.location = location.ToString();
                    match.updated_at = DateTime.Now;
                    match.completed = 0;
                    #region match.body = body ...
                    if (body != null)
                        match.body = body;
                    #endregion
                    #region match.response = response ...
                    if (response != null)
                        match.response = response;
                    #endregion
                    #region match.expectionString = expectionString ...
                    if (response != null)
                        match.expectionString = expectionString;
                    #endregion
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

        public bool                 AppointmentsUpdate(string oldGlobalId, string newGlobalId)
        {
            try
            {
                using (var db = GetContextO())
                {
                    var match = db.appointments.SingleOrDefault(x => x.global_id == oldGlobalId);

                    if (match == null)
                        return false;

                    match.global_id = newGlobalId;
                    match.updated_at = DateTime.Now;
                    match.completed = 0;
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

                    short? temp = match.completed;

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

        public bool                 AppointmentsDelete(int id)
        {
            try
            {
                using (var db = GetContextO())
                {
                    //WARNING - not like others

                    var match = db.appointments.SingleOrDefault(x => x.id == id);

                    if (match == null)
                        return false;

                    match.updated_at = DateTime.Now;
                    match.workflow_state = "Removed";
                    match.version = match.version + 1;

                    db.appointment_versions.Add(MapAppointmentVersions(match));
                    db.SaveChanges();

                    db.appointments.Remove(match);
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
        #endregion

        #region public Lookup
        public bool                 LookupCreateUpdate(string title, string value)
        {
            try
            {
                using (var db = GetContextO())
                {
                    if (string.IsNullOrEmpty(title))
                        return false;
                    if (string.IsNullOrEmpty(value))
                        return false;

                    lookups match = db.lookups.SingleOrDefault(x => x.title == title);

                    if (match == null)
                    {
                        match = new lookups();
                        match.title = title;
                        match.value = value;

                        db.lookups.Add(match);
                        db.SaveChanges();

                        return true;
                    }
                    else
                    {
                        match.value = value;

                        db.SaveChanges();

                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                log.LogEverything("Not Specified", t.PrintException(t.GetMethodName() + " failed, for title:'" + title + "'", ex));
                return false;
            }
        }

        public string               LookupRead(string title)
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

        public List<lookups>        LookupReadAll()
        {
            try
            {
                using (var db = GetContextO())
                {
                    List<lookups> lst = db.lookups.ToList();
                    return lst;
                }
            }
            catch (Exception ex)
            {
                log.LogEverything("Not Specified", t.PrintException(t.GetMethodName() + " failed", ex));
                return null;
            }
        }
        
        public bool                 LookupDelete(string title)
        {
            try
            {
                using (var db = GetContextO())
                {
                    lookups match = db.lookups.Single(x => x.title == title);
               
                    if (match != null)
                    {
                        db.lookups.Remove(match);
                        db.SaveChanges();

                        return true;
                    }
                    else
                        return false;
                }
            }
            catch (Exception ex)
            {
                log.LogEverything("Not Specified", t.PrintException(t.GetMethodName() + " failed, for title:'" + title + "'", ex));
                return false;
            }
        }
        #endregion

        #region public SDK
        public bool                 SyncInteractionCase()
        {
            // read input
            #region create
            appointments appoint = AppointmentsFindOne(LocationOptions.Processed);

            if (appoint != null)
            {
                if (InteractionCaseCreate(appoint))
                {
                    log.LogVariable("Not Specified", nameof(appoint), appoint.ToString());
                    log.LogStandard("Not Specified", "Appointment created in SDK input");
                    return true;
                }
                else
                {
                    log.LogVariable("Not Specified", nameof(appoint), appoint.ToString());
                    log.LogException("Not Specified", "Failed to created Appointment in SDK input", new Exception("FATAL issue"), true);
                    return false;
                }
            }
            #endregion

            #region delete
            appoint = AppointmentsFindOne(LocationOptions.Canceled);

            if (appoint != null)
            {
                if (InteractionCaseDelete(appoint))
                {
                    bool isUpdated = AppointmentsUpdate(appoint.global_id, LocationOptions.Revoked, appoint.body, appoint.expectionString, null);

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

                    int interCaseId = sdkSqlCon.InteractionCaseCreate((int)appointment.template_id, "", siteIds, appointment.global_id, t.Bool(appointment.connected), replacements);

                    var match = db.appointments.Single(x => x.global_id == appointment.global_id);

                    match.location = "Created";
                    match.completed = 0;
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
                AppointmentsUpdate(appointment.global_id, LocationOptions.Exception, appointment.body, t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex), null);
                return false;
            }
        }

        public bool                 InteractionCaseDelete(appointments appointment)
        {
            try
            {
                string mUID = appointment.microting_uid;

                if (string.IsNullOrEmpty(mUID))
                    return true;

                sdkSqlCon.InteractionCaseDelete(int.Parse(mUID));
                return true;
            }
            catch (Exception ex)
            {
                log.LogWarning("Not Specified", t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex));
                AppointmentsUpdate(appointment.global_id, LocationOptions.Exception, appointment.body, t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex), null);
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
                    List<string> lstSent = new List<string>();
                    List<string> lstRetrived = new List<string>();
                    List<string> lstCompleted = new List<string>();
                    List<string> lstDeleted = new List<string>();
                    List<string> lstExpection = new List<string>();
                    bool flagException = false;
                    bool anyCompleted = false;
                    #endregion
                    foreach (var item in match.a_interaction_case_lists)
                    {
                        #region if stat ...
                        statCur = 0;

                        if (item.stat == "Created")
                            statCur = 1;
                        if (item.stat == "Sent")
                        {
                            statCur = 2;
                                 lstSent.Add(item.updated_at + " / " + SiteLookupName(item.siteId) + "     (http://angular/case/" + item.siteId + "/" + item.microting_uid + ")");
                        }
                        if (item.stat == "Retrived")
                        {
                            statCur = 3;
                             lstRetrived.Add(item.updated_at + " / " + SiteLookupName(item.siteId) + "     (http://angular/case/" + item.siteId + "/" + item.microting_uid + ")");
                        }
                        if (item.stat == "Completed")
                        {
                            statCur = 4;
                            anyCompleted = true;
                            lstCompleted.Add(item.updated_at + " / " + SiteLookupName(item.siteId) + "     (http://angular/case/" + item.siteId + "/" + item.microting_uid + ")");
                        }
                        if (item.stat == "Deleted")
                        {
                            statCur = 5;
                              lstDeleted.Add(item.updated_at + " / " + SiteLookupName(item.siteId) + "     (http://angular/case/" + item.siteId + "/" + item.microting_uid + ")");
                        }

                        if (item.stat == "Expection")
                        {
                            flagException = true;
                            lstExpection.Add(item.updated_at + " / " + SiteLookupName(item.siteId) + "     (http://angular/case/" + item.siteId + "/" + item.microting_uid + ")");
                        }

                        if (statHigh < statCur)
                            statHigh = statCur;

                        if (statLow > statCur)
                            statLow = statCur;
                        #endregion
                    }

                    #region pick color
                    if (anyCompleted && statHigh == 5) //as in 1 or more completed, and some deleted
                        statHigh = 4;

                    if (match.workflow_state == "failed to sync")
                        flagException = true;

                    if (t.Bool(AppointmentsFind(match.custom).color_rule))
                        statFinal = statHigh;
                    else
                        statFinal = statLow;
                    #endregion

                    #region craft body text to be added
                    if (lstExpection.Count > 0)
                    {
                        addToBody += "Expection:" + Environment.NewLine;
                        foreach (var line in lstExpection)
                            addToBody += line + Environment.NewLine;
                        addToBody += Environment.NewLine;
                    }

                    if (lstCompleted.Count > 0)
                    {
                        addToBody += "Completed:" + Environment.NewLine;
                        foreach (var line in lstCompleted)
                            addToBody += line + Environment.NewLine;
                        addToBody += Environment.NewLine;
                    }

                    if (lstRetrived.Count > 0)
                    {
                        addToBody += "Retrived:" + Environment.NewLine;
                        foreach (var line in lstRetrived)
                            addToBody += line + Environment.NewLine;
                        addToBody += Environment.NewLine;
                    }

                    if (lstSent.Count > 0)
                    {
                        addToBody += "Sent:" + Environment.NewLine;
                        foreach (var line in lstSent)
                            addToBody += line + Environment.NewLine;
                        addToBody += Environment.NewLine;
                    }

                    if (lstDeleted.Count > 0)
                    {
                        addToBody += "Deleted:" + Environment.NewLine;
                        foreach (var line in lstDeleted)
                            addToBody += line + Environment.NewLine;
                        addToBody += Environment.NewLine;
                    }
                    #endregion

                    #region WorkflowState wFS = ...
                    LocationOptions wFS = LocationOptions.Failed_to_intrepid;
                    if (statFinal == 1)
                        wFS = LocationOptions.Created;
                    if (statFinal == 2)
                        wFS = LocationOptions.Sent;
                    if (statFinal == 3)
                        wFS = LocationOptions.Retrived;
                    if (statFinal == 4)
                        wFS = LocationOptions.Completed;
                    if (statFinal == 5)
                        wFS = LocationOptions.Revoked;
                    if (flagException == true)
                        wFS = LocationOptions.Failed_to_intrepid;
                    #endregion

                    if (addToBody != "")
                        AppointmentsUpdate(match.custom, wFS, null, match.expectionString, addToBody.Trim());
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

        public string               SiteLookupName(int? siteUId)
        {
            try
            {
                if (siteUId == null)
                    return "'Null'";

                var site = sdkSqlCon.SiteRead((int)siteUId);

                if (site == null)
                    return "No matching name found";
                else
                    return site.SiteName;
            }
            catch (Exception ex)
            {
                log.LogWarning("Not Specified", t.PrintException(t.GetMethodName() + " failed to create, for the following reason:", ex));
                return "No matching name found";
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
            SettingCreate(Settings.responseBeforeBody);
            SettingCreate(Settings.calendarName);
            SettingCreate(Settings.userEmailAddress);

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
                    case Settings.responseBeforeBody:       id = 11;    defaultValue = "false";                                 break;
                    case Settings.calendarName:             id = 12;    defaultValue = "Calendar";                              break;
                    case Settings.userEmailAddress:          id = 13;    defaultValue = "no-reply@invalid.invalid";              break;

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
                if (log == null)
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

        //private

        #region mappers
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
        public bool                 UnitTest_TruncateTable(string tableName)
        {
            try
            {
                using (var db = GetContextO())
                {
                    if (msSql)
                    {
                        db.Database.ExecuteSqlCommand("DELETE FROM [dbo].[" + tableName + "];");
                        db.Database.ExecuteSqlCommand("DBCC CHECKIDENT('" + tableName + "', RESEED, 1);");

                        return true;
                    }
                    else
                    {
                        db.Database.ExecuteSqlCommand("SET FOREIGN_KEY_CHECKS=0");
                        db.Database.ExecuteSqlCommand("TRUNCATE TABLE " + tableName + ";");
                        db.Database.ExecuteSqlCommand("SET FOREIGN_KEY_CHECKS=1");

                        return true;
                    }
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
                    if (msSql)
                    {
                        db.Database.ExecuteSqlCommand("DELETE FROM [dbo].[" + tableName + "];");
                        db.Database.ExecuteSqlCommand("DBCC CHECKIDENT('" + tableName + "', RESEED, 1);");

                        return true;
                    }
                    else
                    {
                        db.Database.ExecuteSqlCommand("SET FOREIGN_KEY_CHECKS=0");
                        db.Database.ExecuteSqlCommand("TRUNCATE TABLE " + tableName + ";");
                        db.Database.ExecuteSqlCommand("SET FOREIGN_KEY_CHECKS=1");

                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                string str = ex.Message;
                return false;
            }
        }

        public bool                 UnitTest_OutlookDatabaseClear()
        {
            try
            {
                using (var db = GetContextM())
                {
                    UnitTest_TruncateTable(typeof(appointment_versions).Name);
                    UnitTest_TruncateTable(typeof(appointments).Name);
                    UnitTest_TruncateTable(typeof(lookups).Name);

                    return true;
                }
            }
            catch (Exception ex)
            {
                string str = ex.Message;
                return false;
            }
        }

        public int                  UnitTest_FindLog(int checkCount, string checkValue)
        {
            try
            {
                using (var db = GetContextO())
                {
                    List<logs> lst = db.logs.OrderByDescending(x => x.id).Take(checkCount).ToList();
                    int count = 0;

                    foreach (logs item in lst)
                        if (item.message.Contains(checkValue))
                            count++;

                    return count;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("UnitTest_FindAllActiveEntities failed", ex);
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
        responseBeforeBody,
        calendarName,
        userEmailAddress
    }
}