using Microting.OutlookAddon.Messages;
using OutlookSql;
using OutlookOfficeOnline;
using Rebus.Handlers;
using System.Linq;
using System.Threading.Tasks;

namespace Microting.OutlookAddon.Handlers
{
    public class EformRetrievedHandler : IHandleMessages<EformRetrieved>
    {
        private readonly SqlController sqlController;
        private readonly IOutlookOnlineController outlookOnlineController;

        public EformRetrievedHandler(SqlController sqlController, IOutlookOnlineController outlookOnlineController)
        {
            this.sqlController = sqlController;
            this.outlookOnlineController = outlookOnlineController;
        }

#pragma warning disable 1998
        public async Task Handle(EformRetrieved message)
        {
            Appointment appo = sqlController.AppointmentFindByCaseId(message.caseId);
            outlookOnlineController.CalendarItemUpdate(appo.GlobalId, appo.Start, ProcessingStateOptions.Retrived, appo.Body);
            sqlController.AppointmentsUpdate(appo.GlobalId, ProcessingStateOptions.Retrived, appo.Body, "", "", true, appo.Start, appo.End, appo.Duration);
            sqlController.AppointmentSiteUpdate((int)appo.AppointmentSites.First().Id, message.caseId, ProcessingStateOptions.Retrived);

        }
    }
}