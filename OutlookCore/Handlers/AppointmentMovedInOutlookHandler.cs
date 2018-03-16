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
    public class AppointmentMovedInOutlookHandler : IHandleMessages<AppointmentMovedInOutlook>
    {
        private readonly SqlController sqlController;

        public AppointmentMovedInOutlookHandler(SqlController sqlController)
        {
            this.sqlController = sqlController;
        }

#pragma warning disable 1998
        public async Task Handle(AppointmentMovedInOutlook message)
        {
        }
    }
}
