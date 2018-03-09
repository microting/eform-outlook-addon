using Microting.OutlookAddon.Messages;
using OutlookSql;
using Rebus.Handlers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microting.OutlookAddon.Handlers
{
    class AppointmentCreatedInOutlookHandler : IHandleMessages<AppointmentCreatedInOutlook>
    {
        private readonly SqlController sqlController;

        public AppointmentCreatedInOutlookHandler(SqlController sqlController)
        {
            this.sqlController = sqlController;
        }

#pragma warning disable 1998
        public async Task Handle(AppointmentCreatedInOutlook message)
        {
        }
    }
}
