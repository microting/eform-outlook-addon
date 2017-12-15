using eFormShared;
using OutlookSql;
using Outlook = Microsoft.Office.Interop.Outlook;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OutlookExchangeOnlineAPI;

namespace OutlookOfficeOnline
{
    public class OutlookOnlineController : IOutlookOnlineController
    {
        #region var
        string calendarName;
        SqlController sqlController;
        Log log;
        Tools t = new Tools();
        object _lockOutlook = new object();

        OutlookExchangeOnlineAPIClient outlookExchangeOnlineAPIClient;
        string userEmailAddess;
        #endregion

        #region con
        public OutlookOnlineController(SqlController sqlController, Log log, OutlookExchangeOnlineAPIClient outlookExchangeOnlineAPIClient)
        {
            this.sqlController = sqlController;
            this.log = log;
            this.outlookExchangeOnlineAPIClient = outlookExchangeOnlineAPIClient;
        }
        #endregion

        #region public
        public bool CalendarItemConvertRecurrences()
        {
            //return false;
            try
            {
                bool ConvertedAny = false;
                #region var
                DateTime checkLast_At = DateTime.Parse(sqlController.SettingRead(Settings.checkLast_At));
                double checkPreSend_Hours = double.Parse(sqlController.SettingRead(Settings.checkPreSend_Hours));
                double checkRetrace_Hours = double.Parse(sqlController.SettingRead(Settings.checkRetrace_Hours));
                int checkEvery_Mins = int.Parse(sqlController.SettingRead(Settings.checkEvery_Mins));
                bool includeBlankLocations = bool.Parse(sqlController.SettingRead(Settings.includeBlankLocations));

                DateTime timeOfRun = DateTime.Now;
                DateTime tLimitTo = timeOfRun.AddHours(+checkPreSend_Hours);
                DateTime tLimitFrom = checkLast_At.AddHours(-checkRetrace_Hours);
                #endregion

                #region convert recurrences
                List<Event> eventList = GetCalendarItems(tLimitFrom, tLimitTo);
                if (eventList != null) {
                    foreach (Event item in eventList)
                    {
                        if (item.Type.Equals("Occurrence")) //is recurring, otherwise ignore
                        {
                            #region location "planned"?
                            string location = null;
                            try {
                                location = item.Location.DisplayName;
                            } catch (Exception ex)
                            {
                                log.LogEverything("Not Specified", t.GetMethodName() + " got exception :" + ex.Message + " when trying to do item.Location.DisplayName for item with id : " + item.Id);
                                return false;
                            }
                             

                            if (string.IsNullOrEmpty(location))
                            {
                                if (includeBlankLocations)
                                    location = "planned";
                                else
                                    location = "";
                            }

                            location = location.ToLower();
                            #endregion

                            if (location == "planned")
                            #region ...
                            {
                                try
                                {
                                    string appointmendId = CalendarItemCreate(location, item.Start.DateTime, (item.End.DateTime - item.Start.DateTime).Minutes, item.Subject,
                                    item.BodyPreview);
                                } catch (Exception ex)
                                {
                                    log.LogEverything("Not Specified", t.GetMethodName() + " got exception :" + ex.Message + " when trying to do CalendarItemCreate for item with id : " + item.Id);
                                    return false;
                                }
                                
                                if (CalendarItemDelete(item.Id))
                                {
                                    log.LogStandard("Not Specified", item.Id + " / " + item.Start.DateTime + " converted to non-recurence appointment");
                                    ConvertedAny = true;
                                }
                                
                            }
                            #endregion
                        }
                    }
                } else { }
                
                #endregion

                if (ConvertedAny)
                    log.LogStandard("Not Specified", t.GetMethodName() + " completed + converted appointment(s)");
                else
                    log.LogEverything("Not Specified", t.GetMethodName() + " completed");

                return ConvertedAny;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public bool ParseCalendarItems()
        {
            try
            {
                log.LogStandard("Not Specified", "ParseCalendarItems called. Before setting all vars");
                bool AllParsed = false;
                #region var
                DateTime checkLast_At = DateTime.Parse(sqlController.SettingRead(Settings.checkLast_At));
                double checkPreSend_Hours = double.Parse(sqlController.SettingRead(Settings.checkPreSend_Hours));
                double checkRetrace_Hours = double.Parse(sqlController.SettingRead(Settings.checkRetrace_Hours));
                int checkEvery_Mins = int.Parse(sqlController.SettingRead(Settings.checkEvery_Mins));
                bool includeBlankLocations = bool.Parse(sqlController.SettingRead(Settings.includeBlankLocations));

                DateTime timeOfRun = DateTime.Now;
                DateTime tLimitTo = timeOfRun.AddHours(+checkPreSend_Hours);
                DateTime tLimitFrom = checkLast_At.AddHours(-checkRetrace_Hours);
                #endregion

                log.LogStandard("Not Specified", "ParseCalendarItems called. After setting all vars");

                #region process appointments
                List<Event> eventList = GetCalendarItems(tLimitFrom, tLimitTo);
                if (eventList == null)
                {
                    AllParsed = true;
                }
                else
                {
                    foreach (Event item in eventList)
                    {
                        if (item.Type == "SingleInstance") //is NOT recurring, otherwise ignore
                        {
                            if (tLimitFrom <= item.Start.DateTime && item.Start.DateTime <= tLimitTo)
                            {
                                #region location "planned"?
                                string location = null;
                                try
                                {
                                    location = item.Location.DisplayName;
                                } catch { }
                                

                                if (string.IsNullOrEmpty(location))
                                {
                                    if (includeBlankLocations)
                                        location = "planned";
                                    else
                                        location = "";
                                }

                                location = location.ToLower();
                                #endregion

                                if (location.ToLower() == "planned")
                                #region ...
                                {
                                    log.LogVariable("Not Specified", nameof(location), location);

                                    if (item.BodyPreview != null)
                                        if (item.BodyPreview.Contains("<<< "))
                                            if (item.BodyPreview.Contains("End >>>"))
                                            {
                                                item.BodyPreview = t.ReplaceAtLocationAll(item.BodyPreview, "<<< ", "End >>>", "", true);
                                                item.BodyPreview = item.BodyPreview.Replace("<<< End >>>", "");
                                                item.BodyPreview = item.BodyPreview.Trim();
                                            }

                                    log.LogStandard("Not Specified", "Trying to do UpdateEvent on item.Id:" + item.Id + " to have new location location : " + location);
                                    Event updatedItem = outlookExchangeOnlineAPIClient.UpdateEvent(userEmailAddess, item.Id, "{\"Location\": {\"DisplayName\": \"" + location + "\"},\"Body\": {\"ContentType\": \"HTML\",\"Content\": \"" + ReplaceLinesInBody(item.BodyPreview) + "\"}}");

                                    if (updatedItem == null)
                                    {
                                        return false;
                                    }

                                    log.LogStandard("Not Specified", "Trying create new appointment for item.Id : " + item.Id + " and the UpdateEvent returned Updateditem: " + updatedItem.ToString());

                                    Appointment appo = new Appointment(item.Id, item.Start.DateTime, (item.End.DateTime - item.Start.DateTime).Minutes, item.Subject, updatedItem.Location.DisplayName, updatedItem.BodyPreview, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true);

                                    if (appo.Location == null)
                                        appo.Location = "planned";

                                    if (appo.Location.ToLower() == "planned")
                                    {
                                        log.LogStandard("Not Specified", "Before calling CalendarItemIntrepret.AppointmentsCreate");
                                        int count = sqlController.AppointmentsCreate(appo);
                                        log.LogStandard("Not Specified", "After calling CalendarItemIntrepret.AppointmentsCreate");

                                        if (count > 0)
                                        {
                                            log.LogStandard("Not Specified", "Appointment created successfully for item.Id : " + item.Id);
                                            CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Processed, appo.Body);
                                        }                                          
                                        else
                                        {
                                            if (count == 0)
                                            {
                                                CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Exception, appo.Body);
                                            }                                             
                                            if (count == -1)
                                            {
                                                log.LogStandard("Not Specified", "Appointment not created successfully for item.Id : " + item.Id);

                                                #region appo.Body = 'text'
                                                appo.Body = "<<< Intrepid error: Start >>>" +
                                                    Environment.NewLine + "Global ID already exists in the database." +
                                                    Environment.NewLine + "Indicating that this appointment has already been created." +
                                                    Environment.NewLine + "Likely course, is that you set the Appointment’s location to 'planned'/[blank] again." +
                                                    Environment.NewLine + "" +
                                                    Environment.NewLine + "If you wanted to a create a new appointment in the calendar:" +
                                                    Environment.NewLine + "- Create a new appointment in the calendar" +
                                                    Environment.NewLine + "- Create or copy the wanted details to the new appointment" +
                                                    Environment.NewLine + "" +
                                                    Environment.NewLine + "If you want to restore this appointment’s correct status:" +
                                                    Environment.NewLine + "- Set the appointment’s location to 'check'" +
                                                    Environment.NewLine + "Item.Id :" + item.Id +
                                                    Environment.NewLine + "<<< Intrepid error: End >>>" +
                                                    Environment.NewLine + "" +
                                                    Environment.NewLine + appo.Body;
                                                #endregion
                                                CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Failed_to_intrepret, appo.Body);
                                            }
                                        }
                                    }
                                    else
                                        CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Failed_to_intrepret, appo.Body);

                                    AllParsed = true;
                                }
                                #endregion

                                if (location.ToLower() == "cancel")
                                #region ...
                                {
                                    log.LogVariable("Not Specified", nameof(location), location);

                                    Appointment appo = new Appointment(item.Id, item.Start.DateTime, (item.End.DateTime - item.Start.DateTime).Minutes, item.Subject, item.Location.DisplayName, ReplaceLinesInBody(item.BodyPreview), t.Bool(sqlController.SettingRead(Settings.colorsRule)), true);

                                    if (sqlController.AppointmentsCancel(appo.GlobalId))
                                        CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Canceled, appo.Body);
                                    else
                                        CalendarItemUpdate(appo.GlobalId, appo.Start, LocationOptions.Failed_to_intrepret, appo.Body);

                                    AllParsed = true;
                                }
                                #endregion

                                if (location.ToLower() == "check")
                                #region ...
                                {
                                    log.LogVariable("Not Specified", nameof(location), location);

                                    eFormSqlController.SqlController sqlMicroting = new eFormSqlController.SqlController(sqlController.SettingRead(Settings.microtingDb));
                                    eFormCommunicator.Communicator com = new eFormCommunicator.Communicator(sqlMicroting, log);

                                    var temp = sqlController.AppointmentsFind(item.Id);

                                    var list = sqlMicroting.InteractionCaseListRead(int.Parse(temp.microting_uuid));
                                    foreach (var aCase in list)
                                        com.CheckStatusUpdateIfNeeded(aCase.microting_uid);

                                    CalendarItemReflecting(item.Id);
                                    AllParsed = true;
                                }
                                #endregion
                            }
                        }
                    }
                }
                #endregion

                sqlController.SettingUpdate(Settings.checkLast_At, timeOfRun.ToString());
                log.LogVariable("Not Specified", nameof(Settings.checkLast_At), timeOfRun.ToString());

                return AllParsed;
            }
            catch (Exception ex)
            {
                var lineNumber = 0;
                const string lineSearch = ":line ";
                var index = ex.StackTrace.LastIndexOf(lineSearch);
                if (index != -1)
                {
                    var lineNumerText = ex.StackTrace.Substring(index + lineSearch.Length);
                    if (int.TryParse(lineNumerText, out lineNumber))
                    {
                    }
                }
                log.LogException("Exception", ex.Message + " Exception at line" + lineNumber.ToString(), ex, false);
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public bool CalendarItemReflecting(string globalId)
        {
            try
            {
                #region appointment = 'find one';
                appointments appointment = null;
                string Categories = null;
                if (globalId == null)
                    appointment = sqlController.AppointmentsFindOne(0);
                else
                    appointment = sqlController.AppointmentsFind(globalId);

                if (appointment == null) //double checks status if no new found
                    appointment = sqlController.AppointmentsFindOne(1);
                #endregion

                if (appointment == null)
                    return false;
                log.LogVariable("Not Specified", nameof(appointments), appointment.ToString());

                Event item = AppointmentItemFind(appointment.global_id, appointment.start_at.Value.AddHours(-36), appointment.start_at.Value.AddHours(36)); // TODO!
                if (item != null)
                {
                    item.Location.DisplayName = appointment.processing_state;
                    #region item.Categories = 'workflowState'...
                    switch (appointment.processing_state)
                    {
                        case "Planned":
                            Categories = null;
                            break;
                        case "Processed":
                            Categories = CalendarItemCategory.Processing.ToString();
                            break;
                        case "Created":
                            Categories = CalendarItemCategory.Processing.ToString();
                            break;
                        case "Sent":
                            Categories = CalendarItemCategory.Sent.ToString();
                            break;
                        case "Retrived":
                            Categories = CalendarItemCategory.Retrived.ToString();
                            break;
                        case "Completed":
                            Categories = CalendarItemCategory.Completed.ToString();
                            break;
                        case "Canceled":
                            Categories = CalendarItemCategory.Revoked.ToString();
                            break;
                        case "Revoked":
                            Categories = CalendarItemCategory.Revoked.ToString();
                            break;
                        case "Exception":
                            Categories = CalendarItemCategory.Error.ToString();
                            break;
                        case "Failed_to_intrepid":
                            Categories = CalendarItemCategory.Error.ToString();
                            break;
                        default:
                            Categories = CalendarItemCategory.Error.ToString();
                            break;
                    }
                    #endregion
                    item.BodyPreview = appointment.body;
                    #region item.Body = appointment.expectionString + item.Body + appointment.response ...
                    if (!string.IsNullOrEmpty(appointment.response))
                    {
                        if (t.Bool(sqlController.SettingRead(Settings.responseBeforeBody)))
                        {
                            item.BodyPreview = "<<< Response: Start >>>" +
                            Environment.NewLine +
                            Environment.NewLine + appointment.response +
                            Environment.NewLine +
                            Environment.NewLine + "<<< Response: End >>>" +
                            Environment.NewLine +
                            Environment.NewLine + item.BodyPreview;
                        }
                        else
                        {
                            item.BodyPreview = item.BodyPreview +
                            Environment.NewLine +
                            Environment.NewLine + "<<< Response: Start >>>" +
                            Environment.NewLine +
                            Environment.NewLine + appointment.response +
                            Environment.NewLine +
                            Environment.NewLine + "<<< Response: End >>>";
                        }
                    }
                    if (!string.IsNullOrEmpty(appointment.exceptionString))
                    {
                        item.BodyPreview = "<<< Exception: Start >>>" +
                        Environment.NewLine +
                        Environment.NewLine + appointment.exceptionString +
                        Environment.NewLine +
                        Environment.NewLine + "<<< Exception: End >>>" +
                        Environment.NewLine +
                        Environment.NewLine + item.BodyPreview;
                    }
                    #endregion
                    Event eresult = outlookExchangeOnlineAPIClient.UpdateEvent(userEmailAddess, item.Id, CalendarItemUpdateBody(item.BodyPreview, item.Location.DisplayName, Categories));
                    if (eresult == null)
                    {
                        return false;
                    } else
                    {
                        log.LogStandard("Not Specified", "globalId:'" + appointment.global_id + "' reflected in calendar");
                    }
                        
                    
                }
                else
                    log.LogWarning("Not Specified", "globalId:'" + appointment.global_id + "' no longer in calendar, so hence is considered to be reflected in calendar");

                sqlController.AppointmentsReflected(appointment.global_id);
                log.LogStandard("Not Specified", "globalId:'" + appointment.global_id + "' reflected in database");
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public string CalendarItemCreate(string location, DateTime start, int duration, string subject, string body)
        {
            try
            {
                Event newAppo = outlookExchangeOnlineAPIClient.CreateEvent(userEmailAddess, GetCalendarID(), CalendarItemCreateBody(subject, body, location, start, start.AddMinutes(duration)));
                log.LogStandard("Not Specified", "Calendar item created in " + calendarName);
                return newAppo.Id;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }

        public bool CalendarItemUpdate(string globalId, DateTime start, LocationOptions workflowState, string body)
        {
            log.LogStandard("Not Specified", "CalendarItemUpdate incoming start is : " + start.ToString());
            Event item = AppointmentItemFind(globalId, start.AddHours(-36), start.AddHours(36)); // TODO!

            item.BodyPreview = body;
            item.Location.DisplayName = workflowState.ToString();
            string Categories = null;
            #region item.Categories = 'workflowState'...
            switch (workflowState)
            {
                case LocationOptions.Planned:
                    Categories = null;
                    break;
                case LocationOptions.Processed:
                    Categories = CalendarItemCategory.Processing.ToString();
                    break;
                case LocationOptions.Created:
                    Categories = CalendarItemCategory.Processing.ToString();
                    break;
                case LocationOptions.Sent:
                    Categories = CalendarItemCategory.Sent.ToString();
                    break;
                case LocationOptions.Retrived:
                    Categories = CalendarItemCategory.Retrived.ToString();
                    break;
                case LocationOptions.Completed:
                    Categories = CalendarItemCategory.Completed.ToString();
                    break;
                case LocationOptions.Canceled:
                    Categories = CalendarItemCategory.Revoked.ToString();
                    break;
                case LocationOptions.Revoked:
                    Categories = CalendarItemCategory.Revoked.ToString();
                    break;
                case LocationOptions.Exception:
                    Categories = CalendarItemCategory.Error.ToString();
                    break;
                case LocationOptions.Failed_to_intrepret:
                    Categories = CalendarItemCategory.Error.ToString();
                    break;
            }
            #endregion

            Event eresult = outlookExchangeOnlineAPIClient.UpdateEvent(userEmailAddess, item.Id, CalendarItemUpdateBody(item.BodyPreview, item.Location.DisplayName, Categories));
            if (eresult == null)
            {
                log.LogStandard("Not Specified", AppointmentPrint(item) + " NOT updated to " + workflowState.ToString());
                return false;
            } else
            {
                log.LogStandard("Not Specified", AppointmentPrint(item) + " updated to " + workflowState.ToString());
                return true;
            }

        }

        public bool CalendarItemDelete(string globalId)
        {
            log.LogEverything("Not Specified", "OutlookOnlineController.CalendarItemDelete called");
            try
            {
                userEmailAddess = GetUserEmailAddress();
                outlookExchangeOnlineAPIClient.DeleteEvent(userEmailAddess, globalId);
                log.LogStandard("Not Specified", "globalId:'" + globalId + "' deleted");
                return true;
            } catch (Exception ex)
            {
                log.LogStandard("Not Specified", "CalendarItemDelete failed got exception:" + ex.Message);
                return false;
            }
            
        }
        #endregion

        #region private
        private DateTime RoundTime(DateTime dTime)
        {
            dTime = dTime.AddMinutes(1);
            dTime = new DateTime(dTime.Year, dTime.Month, dTime.Day, dTime.Hour, 0, 0);
            log.LogVariable("Not Specified", nameof(dTime), dTime);
            return dTime;
        }
        private Event AppointmentItemFind(string globalId, DateTime tLimitFrom, DateTime tLimitTo)
        {
            try
            {
                log.LogEverything("Not Specified", "OutlookOnlineController.AppointmentItemFind called");
                string filter = "AppointmentItemFind [After] '" + tLimitFrom.ToString("g") + "' AND [before] <= '" + tLimitTo.ToString("g") + "'";
                log.LogVariable("Not Specified", nameof(filter), filter.ToString());

                List<Event> calendarItemsAllToday = GetCalendarItems(tLimitFrom, tLimitTo);
                foreach (Event item in calendarItemsAllToday)
                {
                    //log.LogEverything("Not Specified", "OutlookOnlineController.AppointmentItemFind calendarItemsAllToday current item.Id is " + item.Id);
                    if (item.Id.Equals(globalId))
                        return item;
                }

                //List<Event> calendarItemsRes = GetCalendarItems(new DateTime(1975, 1, 1).Date, DateTime.Today.AddDays(1));
                List<Event> calendarItemsRes = GetCalendarItems(new DateTime(1975, 1, 1).Date, new DateTime(2025, 12, 31).Date);
                foreach (Event item in calendarItemsRes)
                {
                    //log.LogEverything("Not Specified", "OutlookOnlineController.AppointmentItemFind calendarItemsRes current item.Id is " + item.Id);
                    if (item.Id.Equals(globalId))
                        return item;
                }

                log.LogEverything("Not Specified", "No match found for " + nameof(globalId) + ":" + globalId);
                return null;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed. Due to no match found global id:" + globalId, ex);
            }
        }

        private string AppointmentPrint(Event appItem)
        {
            return "GlobalId:" + appItem.Id + " / Start:" + appItem.Start + " / Title:" + appItem.Subject;
        }

        private List<Event> GetCalendarItems(DateTime tLimitFrom, DateTime tLimitTo)
        {
            lock (_lockOutlook)
            {
                log.LogEverything("Not Specified", "OutlookOnlineController.GetCalendarItems called");
                string filter = "GetCalendarItems [After] '" + tLimitTo.ToString("g") + "' AND [before] <= '" + tLimitFrom.ToString("g") + "'";
                log.LogVariable("Not Specified", nameof(filter), filter.ToString());
                calendarName = GetCalendarName();
                userEmailAddess = GetUserEmailAddress();
                CalendarList calendarList = outlookExchangeOnlineAPIClient.GetCalendarList(userEmailAddess, calendarName);
                if (calendarList != null)
                {
                    foreach (Calendar cal in calendarList.value)
                    {
                        log.LogEverything("Not Specified", "GetCalendarItems comparing cal.Name " + cal.Name + " with calendarName " + calendarName);
                        if (cal.Name.Equals(calendarName, StringComparison.OrdinalIgnoreCase))
                        {
                            EventList outlookCalendarItems = outlookExchangeOnlineAPIClient.GetCalendarItems(userEmailAddess, cal.Id, tLimitFrom, tLimitTo);
                            //log.LogVariable("Not Specified", "outlookCalendarItems.Count", outlookCalendarItems.value.Count);
                            return outlookCalendarItems.value;
                        }
                    }
                }                
                return null;

            }
        }

        private string GetCalendarName()
        {
            log.LogEverything("Not Specified", "GetCalendarName called");
            if (!string.IsNullOrEmpty(calendarName))
                return calendarName;
            else
            {
                calendarName = sqlController.SettingRead(Settings.calendarName);
                if (!string.IsNullOrEmpty(calendarName))
                    return calendarName;
                else
                    throw new Exception(t.GetMethodName() + " failed, to get calendarName'");
            }

        }
        private string GetUserEmailAddress()
        {
            log.LogEverything("Not Specified", "GetUserEmailAddress called");
            if (!string.IsNullOrEmpty(userEmailAddess))
                return userEmailAddess;
            else
            {
                userEmailAddess = sqlController.SettingRead(Settings.userEmailAddress);
                if (!string.IsNullOrEmpty(userEmailAddess))
                    return userEmailAddess;
                else
                    throw new Exception(t.GetMethodName() + " failed, to get userEmailAddess");
            }
        }
        private string GetCalendarID()
        {
            calendarName = GetCalendarName();
            userEmailAddess = GetUserEmailAddress();
            CalendarList calendarList = outlookExchangeOnlineAPIClient.GetCalendarList(userEmailAddess, calendarName);
            if (calendarList != null)
            {
                foreach (Calendar cal in calendarList.value)
                {
                    if (cal.Name.Equals(calendarName, StringComparison.OrdinalIgnoreCase))
                    {
                        return cal.Id;
                    }
                }
            }
            return null;                   
        }

        #endregion

        #region unit test
        public List<Appointment> UnitTest_CalendarItemGetAllNonRecurring(DateTime startPoint, DateTime endPoint)
        {
            try
            {
                #region var
                List<Appointment> lstAppoint = new List<Appointment>();
                #endregion
                foreach (Event item in GetCalendarItems(new DateTime(1975, 1, 1).Date, DateTime.Today.AddYears(25))) // getting all events from 1975 to today  + 25 years i.e all possible events
                    if (item.Location != null)
                        if (item.Type.Equals("Occurrence"))
                        {
                            //ignore
                        }
                        else
                        {
                            if (startPoint <= item.Start.DateTime && item.Start.DateTime <= endPoint)
                                lstAppoint.Add(new Appointment(item.Id, item.Start.DateTime, (item.End.DateTime - item.Start.DateTime).Minutes, item.Subject, item.Location.DisplayName, item.BodyPreview, t.Bool(sqlController.SettingRead(Settings.colorsRule)), true));
                        }

                return lstAppoint;
            }
            catch (Exception ex)
            {
                throw new Exception(t.GetMethodName() + " failed", ex);
            }
        }


        public bool UnitTest_ForceException(string exceptionType)
        {
            throw new NotImplementedException();
        }

        private string UnitTest_CalendarBody()
        {
            return
                                            "TempLate# " + "’Besked’"
                    + Environment.NewLine + "Sites# " + "’All’"
                    + Environment.NewLine + "title# " + "Outlook appointment eForm test"
                    + Environment.NewLine + "info# " + "Tekst fra Outlook appointment";
        }
        private string CalendarItemUpdateBody(string BodyContent, string Location, string Category)
        {
            TimeZone localZone = TimeZone.CurrentTimeZone;
            string UpdateBody = "{\"Location\": {\"DisplayName\": \"" + Location + "\"},\"Body\": {\"ContentType\": \"HTML\",\"Content\": \"" + ReplaceLinesInBody(BodyContent) + "\"},\"Categories\": [\"" + Category + "\"]}";
            return UpdateBody;
        }
        private string CalendarItemCreateBody(string Subject, string BodyContent, string Location, DateTime Start, DateTime End)
        {
            TimeZone localZone = TimeZone.CurrentTimeZone;
            string CreateBody = "{\"Subject\": \"" + Subject + "\",\"Body\": {\"ContentType\": \"HTML\",\"Content\": \"" + ReplaceLinesInBody(BodyContent) + "\"},\"Location\": {\"DisplayName\": \"" + Location + "\"},\"Start\": {\"DateTime\": \"" + Start + "\",\"TimeZone\": \"" + localZone.StandardName + "\"},\"End\":{\"DateTime\":\"" + End + "\",\"TimeZone\": \"" + localZone.StandardName + "\"}}";
            return CreateBody;
        }
        private string ReplaceLinesInBody(string BodyPreview)
        {
            return BodyPreview.Replace("\r\n", "<br/>"); ;
        }
        #endregion
    }

    public enum CalendarItemCategory
    {
        Completed,
        Error,
        Processing,
        Retrived,
        Revoked,
        Sent
    }
}