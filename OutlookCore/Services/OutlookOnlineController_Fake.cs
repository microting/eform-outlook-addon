using eFormShared;
using OutlookSql;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookOfficeOnline
{
    //public class OutlookOnlineController_Fake : IOutlookOnlineController
    //{
    //    #region var
    //    SqlController sqlController;
    //    Log log;
    //    Tools t = new Tools();
    //    object _lockOutlook = new object();
    //    Random rndm = new Random();

    //    string forceException = "";
    //    #endregion

    //    #region con
    //    public OutlookOnlineController_Fake(SqlController sqlController, Log log)
    //    {
    //        this.sqlController = sqlController;
    //        this.log = log;
    //    }
    //    #endregion

    //    #region public
    //    public bool CalendarItemConvertRecurrences()
    //    {
    //        log.LogEverything("Unit test", t.GetMethodName() + " called");

    //        int temp = rndm.Next(0, 2);
    //        if (temp == 2)
    //            temp = 1;
    //        bool flag = t.Bool(temp + "");

    //        log.LogVariable("Unit test", nameof(flag), flag);
    //        return flag;
    //    }

    //    public bool ParseCalendarItems()
    //    {
    //        log.LogEverything("Unit test", t.GetMethodName() + " called");

    //        if (forceException != "")
    //        {
    //            string exceptionString = forceException + ". Exception as per request";
    //            forceException = "";
    //            throw new Exception(exceptionString);
    //        }

    //        int temp = rndm.Next(0, 2);
    //        if (temp == 2)
    //            temp = 1;
    //        bool flag = t.Bool(temp + "");

    //        log.LogVariable("Unit test", nameof(flag), flag);
    //        return flag;
    //    }

    //    public bool CalendarItemReflecting(string globalId)
    //    {
    //        log.LogStandard("Unit test", t.GetMethodName() + " called");
    //        log.LogVariable("Unit test", (nameof(globalId)), globalId);

    //        if (globalId == null)
    //        {
    //            int temp = rndm.Next(0, 2);
    //            if (temp == 2)
    //                temp = 1;
    //            bool flag = t.Bool(temp + "");

    //            log.LogVariable("Unit test", nameof(flag), flag);
    //            return flag;
    //        }
    //        if (globalId == "")
    //            return false;
    //        if (globalId == "throw new expection")
    //            throw new Exception(t.GetMethodName() + " failed (Exception as per request)");

    //        return true;
    //    }

    //    public string CalendarItemCreate(string location, DateTime start, int duration, string subject, string body, string originalStartTimeZone, string originalEndTimeZone)
    //    {
    //        string globalId = "Faked GlobalId:" + t.GetRandomInt(8);
    //        sqlController.AppointmentsCreate(new Appointment(globalId, start, duration, subject, location, body, false, false, null));
    //        return globalId;
    //    }

    //    public bool CalendarItemUpdate(string globalId, DateTime start, ProcessingStateOptions workflowState, string body)
    //    {
    //        return true;
    //    }

    //    public bool CalendarItemDelete(string globalId)
    //    {
    //        return true;
    //    }
    //    #endregion

    //    #region private
    //    private Appointment CreateAppointment(Appointment appointment)
    //    {
    //        return new Appointment();
    //    }
    //    #endregion

    //    #region unit test
    //    public List<Appointment> UnitTest_CalendarItemGetAllNonRecurring(DateTime startPoint, DateTime endPoint)
    //    {
    //        return new List<Appointment>();
    //    }

    //    public bool UnitTest_ForceException(string exceptionType)
    //    {
    //        forceException = exceptionType;
    //        return true;
    //    }

    //    private string UnitTest_CalendarBody()
    //    {
    //        return
    //                                        "TempLate# " + "’Besked’"
    //                + Environment.NewLine + "Sites# " + "’All’"
    //                + Environment.NewLine + "title# " + "Outlook appointment eForm test"
    //                + Environment.NewLine + "info# " + "Tekst fra Outlook appointment";
    //    }
    //    #endregion
    //}
}