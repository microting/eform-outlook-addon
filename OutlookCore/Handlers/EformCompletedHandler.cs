using Microting.OutlookAddon.Messages;
using OutlookOfficeOnline;
using OutlookSql;
using Rebus.Handlers;
using System.Linq;
using System.Threading.Tasks;

namespace Microting.OutlookAddon.Handlers
{
    public class EformCompletedHandler : IHandleMessages<EformCompleted>
    {
        private readonly SqlController sqlController;
        private readonly IOutlookOnlineController outlookOnlineController;

        public EformCompletedHandler(SqlController sqlController, IOutlookOnlineController outlookOnlineController)
        {
            this.sqlController = sqlController;
            this.outlookOnlineController = outlookOnlineController;
        }

#pragma warning disable 1998
        public async Task Handle(EformCompleted message)
        {
            Appointment appo = sqlController.AppointmentFindByCaseId(message.caseId);
            outlookOnlineController.CalendarItemUpdate(appo.GlobalId, appo.Start, ProcessingStateOptions.Completed, appo.Body);
            sqlController.AppointmentsUpdate(appo.GlobalId, ProcessingStateOptions.Completed, appo.Body, "", "", true, appo.Start, appo.End, appo.Duration);
            sqlController.AppointmentSiteUpdate((int)appo.AppointmentSites.First().Id, message.caseId, ProcessingStateOptions.Completed);

        }
    }
}