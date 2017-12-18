using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookSql
{
    public class AppoinntmentSite
    {
        #region var/pop
        public int? Id { get; set; }
        public int MicrotingSiteUid { get; set; }
        public string ProcessingState { get; set; }
        public string MicrotingUuId { get; set; }
        #endregion

        #region con
        public AppoinntmentSite()
        {

        }

        public AppoinntmentSite(int? id, int microtingSiteUid, string processingState, string microtingUuid)
        {
            Id = id;
            MicrotingSiteUid = microtingSiteUid;
            ProcessingState = processingState;
            MicrotingUuId = microtingUuid;

        }
        #endregion
    }
}
