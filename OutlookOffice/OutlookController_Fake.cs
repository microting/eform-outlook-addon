using eFormShared;
using OutlookSql;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookOffice
{
    public class OutlookController_Fake : IOutlookController
    {
        #region var
        SqlController sqlController;
        Log log;
        Tools t = new Tools();
        object _lockOutlook = new object();
        Random rndm = new Random();

        string forceException = "";
        #endregion

        #region con
        public                      OutlookController_Fake(SqlController sqlController, Log log)
        {
            this.sqlController = sqlController;
            this.log = log;
        }
        #endregion

        #region public
        public bool                 CalendarItemConvertRecurrences()
        {
            log.LogEverything("Unit test", t.GetMethodName() + " called");

            int temp  = rndm.Next(0, 2);
            if (temp == 2)
                temp = 1;
            bool flag = t.Bool(temp + "");

            log.LogVariable("Unit test", nameof(flag), flag);
            return flag;
        }

        public bool                 CalendarItemIntrepid()
        {
            log.LogEverything("Unit test", t.GetMethodName() + " called");

            if (forceException != "")
            {
                string exceptionString = forceException + ". Exception as per request";
                forceException = "";
                throw new Exception(exceptionString);
            }

            int temp = rndm.Next(0, 2);
            if (temp == 2)
                temp = 1;
            bool flag = t.Bool(temp + "");

            log.LogVariable("Unit test", nameof(flag), flag);
            return flag;
        }

        public bool                 CalendarItemReflecting(string globalId)
        {
            log.LogStandard("Unit test", t.GetMethodName() + " called");
            log.LogVariable("Unit test", (nameof(globalId)), globalId);

            if (globalId == null)
            {
                //Copied
                var appointment = sqlController.AppointmentsFindOne(1);

                if (appointment != null)
                {
                    if (appointment.global_id.Contains("Appointment requested to be created"))
                    {
                        appointment.location = "Created";
                        string newGlobalId = CalendarItemCreate(appointment.location, (DateTime)appointment.start_at, (int)appointment.duration, appointment.subject, appointment.body);
                        log.LogCritical("Not Specified", "CalendarItemCreate successful");
                        log.LogVariable("Not Specified", nameof(newGlobalId), newGlobalId);

                        sqlController.AppointmentsUpdate(appointment.global_id, WorkflowState.Created, null, null, null);
                        sqlController.AppointmentsUpdate(appointment.global_id, newGlobalId);
                        log.LogEverything("Not Specified", "AppointmentsUpdate successful");

                        return true;
                    }
                }
                //Copied

                int temp = rndm.Next(0, 2);
                if (temp == 2)
                    temp = 1;
                bool flag = t.Bool(temp + "");

                log.LogVariable("Unit test", nameof(flag), flag);
                return flag;
            }
            if (globalId == "")
                return false;
            if (globalId == "throw new expection")
                throw new Exception(t.GetMethodName() + " failed (Exception as per request)");

            return true;
        }

        public string               CalendarItemCreate(string location, DateTime start, int duration, string subject, string body)
        {
            return "Faked GlobalId:" + t.GetRandomInt(8);
        }

        public bool                 CalendarItemUpdate(string globalId, DateTime start, WorkflowState workflowState, string body)
        {
             return true;
        }

        public bool                 CalendarItemDelete(string globalId, DateTime start)
        {
            return true;
        }
        #endregion

        #region private
        private Appointment         CreateAppointment(Appointment appointment)
        {
            return new Appointment();
        }
        #endregion

        #region unit test
        public List<Appointment>    UnitTest_CalendarItemGetAllNonRecurring(DateTime startPoint, DateTime endPoint)
        {
            return new List<Appointment>();
        }

        public bool                 UnitTest_ForceException(string exceptionType)
        {
            forceException = exceptionType;
            return true;
        }

        private string              UnitTest_CalendarBody()
        {
            return
                                            "TempLate# "+ "’Besked’"
                    + Environment.NewLine + "Sites# "   + "’All’"
                    + Environment.NewLine + "title# "   + "Outlook appointment eForm test"
                    + Environment.NewLine + "info# "    + "Tekst fra Outlook appointment";
        }
        #endregion
    }
}