using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Microting.OutlookAddon.Messages
{
    public class EformRetrieved
    {
        public string caseId { get; protected set; }

        public EformRetrieved(string caseId)
        {
            this.caseId = caseId;
        }
    }
}
