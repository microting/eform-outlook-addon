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
    public class ConvertRecurringEventToSingleEventHandler : IHandleMessages<ConvertRecurringEventToSingleEvent>
    {
        private readonly SqlController sqlController;

        public ConvertRecurringEventToSingleEventHandler(SqlController sqlController)
        {
            this.sqlController = sqlController;
        }

#pragma warning disable 1998
        public async Task Handle(ConvertRecurringEventToSingleEvent message)
        {
        }
    }
}
