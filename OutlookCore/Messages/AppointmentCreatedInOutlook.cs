using OutlookSql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microting.OutlookAddon.Messages
{

    public class AppointmentCreatedInOutlook
    {
        public Appointment Appo { get; protected set; }

        public AppointmentCreatedInOutlook(Appointment appo)
        {
            Appo = appo ?? throw new ArgumentNullException(nameof(appo));
        }
    }
}
