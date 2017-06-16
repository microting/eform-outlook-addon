using eFormShared;
using OutlookShared;
using OutlookSql;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookMicrotingSdk
{
    public class SdkController
    {
        //#region var
        //SqlController sqlController;
        //Tools t = new Tools();
        //#endregion

        //// con
        //public SdkController(SqlController sqlController)
        //{
        //    this.sqlController = sqlController;
        //}

        //// public
        //public bool SyncInteractionCaseCreate()
        //{
        //    // create in input
        //    appointments appoint = sqlController.AppointmentsFindOne(WorkflowState.Processed);

        //    if (appoint != null)
        //    {
        //        bool isCreated = sqlController.InteractionCaseCreate(appoint);

        //        if (isCreated)
        //        {
        //            bool isUpdated = sqlController.AppointmentsUpdate(appoint.global_id, WorkflowState.Created, appoint.body, appoint.expectionString, null);

        //            if (isUpdated)
        //                return true;
        //            else
        //            {
        //                sqlController.LogVariable(nameof(appoint), appoint.ToString());
        //                sqlController.LogException("Failed to update Outlook appointment, but Appointment created in SDK input", new Exception("FATAL issue"), true);
        //            }
        //        }
        //        else
        //        {
        //            sqlController.LogVariable(nameof(appoint), appoint.ToString());
        //            sqlController.LogException("Failed to created Appointment in SDK input", new Exception("FATAL issue"), true);
        //        }

        //        return false;
        //    }

        //    // read from output
        //    return sqlController.InteractionCaseProcessed();
        //}
    }
}