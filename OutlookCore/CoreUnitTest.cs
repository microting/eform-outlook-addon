using eFormShared;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookCore
{
    public class CoreUnitTest
    {
        #region var
        Core core;
        Tools t = new Tools();
        #endregion

        #region con
        public CoreUnitTest(Core core)
        {
            this.core = core;
            core.UnitTest_SetUnittest();
        }
        #endregion

        public void CaseComplet(string microtingUId, string checkUId, int workerUId, int unitUId)
        {
        }

        public bool CoreDead()
        {
            return core.UnitTest_CoreDead();
        }

        public void Close()
        {
            Thread closeCore
                = new Thread(() => core.Close());
            closeCore.Start();
        }
    }
}