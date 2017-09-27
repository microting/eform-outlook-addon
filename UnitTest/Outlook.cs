using eFormShared;
using OutlookSql;
using OutlookCore;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Xunit;
using System.Threading;

namespace UnitTest
{
    public class TestContext : IDisposable
    {
        //string connectionStringLocal_SDK = "Persist Security Info=True;server=localhost;database=" + "Outlook_UnitTest_" + "Microting"        + ";uid=root;password=1234";
        //string connectionStringLocal_OUT = "Persist Security Info=True;server=localhost;database=" + "Outlook_UnitTest_" + "MicrotingOutlook" + ";uid=root;password=1234";

        string connectionStringLocal_SDK = "Data Source=DESKTOP-7V1APE5\\SQLEXPRESS;Initial Catalog=" + "Outlook_UnitTest_" + "Microting"        + ";Integrated Security=True";
        string connectionStringLocal_OUT = "Data Source=DESKTOP-7V1APE5\\SQLEXPRESS;Initial Catalog=" + "Outlook_UnitTest_" + "MicrotingOutlook" + ";Integrated Security=True";

        #region content
        #region var
        SqlController sqlCon;
        string serverConnectionString_SDK = "";
        string serverConnectionString_OUT = "";
        #endregion

        #region once for all tests - build order
        public TestContext()
        {
            try
            {
                if (Environment.MachineName.ToLower().Contains("testing"))
                {
                    serverConnectionString_SDK = "Persist Security Info=True;server=localhost;database=" + "OutlookUnitTest_" + "Microting"        + ";uid=root;password="; //Uses travis database
                    serverConnectionString_OUT = "Persist Security Info=True;server=localhost;database=" + "OutlookUnitTest_" + "MicrotingOutlook" + ";uid=root;password="; //Uses travis database
                }
                else
                {
                    serverConnectionString_SDK = connectionStringLocal_SDK;
                    serverConnectionString_OUT = connectionStringLocal_OUT;
                }
            }
            catch { }

            try
            {
                var sqlSdk = new eFormSqlController.SqlController(serverConnectionString_SDK);
                var adminT = new eFormCore.AdminTools(serverConnectionString_SDK);

                if (sqlSdk.SettingRead(eFormSqlController.Settings.token) == "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")
                    adminT.DbSetup("unittest");

                sqlCon = new SqlController(serverConnectionString_OUT);
                sqlCon.SettingUpdate(Settings.microtingDb, serverConnectionString_SDK);
                sqlCon.SettingUpdate(Settings.calendarName, "unittest");
            }
            catch (Exception ex)
            {
                string temp = ex.Message;
            }
        }
        #endregion

        #region once for all tests - teardown
        public void Dispose()
        {
            //sqlController.UnitTest_DeleteDb();
        }
        #endregion

        public string GetConnectionStringOutlook()
        {
            return serverConnectionString_OUT;
        }

        public string GetConnectionStringSdk()
        {
            return serverConnectionString_SDK;
        }
        #endregion
    }

    [Collection("Database collection")]
    public class Outlook
    {
        #region var
        eFormCore.Core coreSdk;
        Core coreOut;
        eFormCore.CoreUnitTest core_UT;
        eFormSqlController.SqlController sqlConSdk;
        SqlController sqlConOut;
        eFormCore.AdminTools adminTool;
        Tools t = new Tools();

        object _lockTest = new object();
        object _lockFil = new object();

        int siteId1 = 2001;
        //int siteId2 = 2002;
        //int workerMUId = 666;
        //int unitMUId = 345678;

        string connectionStringOut = "";
        string connectionStringSdk = "";
        #endregion

        #region con
        public Outlook(TestContext testContext)
        {
            connectionStringOut = testContext.GetConnectionStringOutlook();
            connectionStringSdk     = testContext.GetConnectionStringSdk();
        }
        #endregion

        #region prepare and teardown     
        private void TestPrepare(string testName, bool startSdk, bool startOut)
        {
            adminTool = new eFormCore.AdminTools(connectionStringSdk);
            string temp = adminTool.DbClear();
            if (temp != "")
                throw new Exception("CleanUp failed");

            sqlConSdk = new eFormSqlController.SqlController(connectionStringSdk);
            sqlConOut = new                    SqlController(connectionStringOut);

            coreSdk = new eFormCore.Core();
            core_UT = new eFormCore.CoreUnitTest(coreSdk);
            coreOut = new Core();

            coreSdk.HandleNotificationNotFound += EventNotificationNotFound;
            coreSdk.HandleEventException += EventException;

            if (startSdk)
                coreSdk.Start(connectionStringSdk);

            if (startOut)
                coreOut.Start(connectionStringOut);
        }

        private void TestTeardown()
        {
            if (coreSdk != null)
                if (coreSdk.Running())
                    core_UT.Close();

            if (coreOut != null)
                if (coreOut.Running())
                    coreOut.Close();
        }
        #endregion

        #region - test 000x virtal basics
        [Fact]
        public void Test000_Basics_1a_MustAlwaysPass()
        {
            lock (_lockTest)
            {
                //Arrange
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                checkValueB = true;

                //Assert
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        //[Fact]
        //public void Test000_Basics_2a_PrepareAndTeardownTestdata()
        //{
        //    lock (_lockTest)
        //    {
        //        //Arrange
        //        TestPrepare(t.GetMethodName(), true, true);
        //        bool checkValueA = true;
        //        bool checkValueB = false;

        //        //Act
        //        checkValueB = true;

        //        //Assert
        //        TestTeardown();
        //        Assert.Equal(checkValueA, checkValueB);
        //    }
        //}

        [Fact]
        public void Test000_Basics_2b_PrepareAndTeardownTestdata()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), true, false);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                checkValueB = true;

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        //[Fact]
        //public void Test000_Basics_2c_PrepareAndTeardownTestdata()
        //{
        //    lock (_lockTest)
        //    {
        //        //Arrange
        //        TestPrepare(t.GetMethodName(), false, true);
        //        bool checkValueA = true;
        //        bool checkValueB = false;

        //        //Act
        //        checkValueB = true;

        //        //Assert
        //        TestTeardown();
        //        Assert.Equal(checkValueA, checkValueB);
        //    }
        //}

        [Fact]
        public void Test000_Basics_2d_PrepareAndTeardownTestdata()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                checkValueB = true;

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }
        #endregion

        #region - test 001x core
        //[Fact]
        //public void Test001_Core_1a_Start_WithNullExpection()
        //{
        //    lock (_lockTest)
        //    {
        //        //Arrange
        //        TestPrepare(t.GetMethodName(), false, false);
        //        string checkValueA = "serverConnectionString is not allowed to be null or empty";
        //        string checkValueB = "";
        //        Core core = new Core();

        //        //Act
        //        try
        //        {
        //            checkValueB = core.Start(null) + "";
        //        }
        //        catch (Exception ex)
        //        {
        //            checkValueB = ex.InnerException.Message;
        //        }

        //        //Assert
        //        TestTeardown();
        //        Assert.Equal(checkValueA, checkValueB);
        //    }
        //}

        //[Fact]
        //public void Test001_Core_1b_Start_WithBlankExpection()
        //{
        //    lock (_lockTest)
        //    {
        //        //Arrange
        //        TestPrepare(t.GetMethodName(), false, false);
        //        string checkValueA = "serverConnectionString is not allowed to be null or empty";
        //        string checkValueB = "";
        //        Core core = new Core();

        //        //Act
        //        try
        //        {
        //            checkValueB = core.Start("").ToString();
        //        }
        //        catch (Exception ex)
        //        {
        //            checkValueB = ex.InnerException.Message;
        //        }

        //        //Assert
        //        TestTeardown();
        //        Assert.Equal(checkValueA, checkValueB);
        //    }
        //}

        //[Fact]
        //public void Test001_Core_3a_Start()
        //{
        //    //Arrange
        //    TestPrepare(t.GetMethodName(), false, false);
        //    string checkValueA = "True";
        //    string checkValueB = "";
        //    Core core = new Core();

        //    //Act
        //    try
        //    {
        //        checkValueB = core.Start(connectionStringOut).ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        checkValueB = t.PrintException(t.GetMethodName() + " failed", ex);
        //    }

        //    //Assert
        //    TestTeardown();
        //    Assert.Equal(checkValueA, checkValueB);
        //}

        //[Fact]
        //public void Test001_Core_4a_IsRunning()
        //{
        //    //Arrange
        //    TestPrepare(t.GetMethodName(), false, false);
        //    string checkValueA = "FalseTrue";
        //    string checkValueB = "";
        //    Core core = new Core();

        //    //Act
        //    try
        //    {
        //        checkValueB = core.Running().ToString();
        //        core.Start(connectionStringOut);
        //        checkValueB += core.Running().ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        checkValueB = t.PrintException(t.GetMethodName() + " failed", ex);
        //    }

        //    //Assert
        //    TestTeardown();
        //    Assert.Equal(checkValueA, checkValueB);
        //}

        //[Fact]
        //public void Test001_Core_5a_Close()
        //{
        //    //Arrange
        //    TestPrepare(t.GetMethodName(), false, false);
        //    string checkValueA = "True";
        //    string checkValueB = "";
        //    Core core = new Core();

        //    //Act
        //    try
        //    {
        //        checkValueB = core.Close().ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        checkValueB = t.PrintException(t.GetMethodName() + " failed", ex);
        //    }

        //    //Assert
        //    TestTeardown();
        //    Assert.Equal(checkValueA, checkValueB);
        //}

        //[Fact]
        //public void Test001_Core_6a_RunningForWhileThenClose()
        //{
        //    //Arrange
        //    TestPrepare(t.GetMethodName(), false, false);
        //    string checkValueA = "FalseTrueTrue";
        //    string checkValueB = "";
        //    Core core = new Core();

        //    //Act
        //    try
        //    {
        //        checkValueB = core.Running().ToString();
        //        core.Start(connectionStringOut);
        //        Thread.Sleep(30000);
        //        checkValueB += core.Running().ToString();
        //        Thread.Sleep(05000);
        //        checkValueB += core.Close().ToString();
        //    }
        //    catch (Exception ex)
        //    {
        //        checkValueB = t.PrintException(t.GetMethodName() + " failed", ex);
        //    }

        //    //Assert
        //    TestTeardown();
        //    Assert.Equal(checkValueA, checkValueB);
        //}
        #endregion

        #region - test 002x - sqlController
        //[Fact]
        //public void Test002_SqlController_1a_TemplateCreateAndRead()
        //{
        //    lock (_lockTest)
        //    {
        //        //Arrange
        //        TestPrepare(t.GetMethodName(), false, false);
        //        appointments checkValueA = null;
        //        appointments checkValueB = new appointments();

        //        //Act
        //        sqlConOut.AppointmentsCreate(null);

        //        checkValueB = sqlConOut.AppointmentsFind(null);

        //        //Assert
        //        TestTeardown();
        //        Assert.Equal(checkValueA, checkValueB);
        //    }
        //}

        //[Fact]
        //public void Test002_SqlController_2a_TemplateCreateAndRead()
        //{
        //    lock (_lockTest)
        //    {
        //        //Arrange
        //        TestPrepare(t.GetMethodName(), false, false);
        //        appointments checkValueA = null;
        //        appointments checkValueB = new appointments();

        //        //Act
        //        checkValueB = sqlConOut.AppointmentsFind(null);

        //        //Assert
        //        TestTeardown();
        //        Assert.Equal(checkValueA, checkValueB);
        //    }
        //}

        //[Fact]
        //public void Test002_SqlController_2b_TemplateCreateAndRead()
        //{
        //    lock (_lockTest)
        //    {
        //        //Arrange
        //        TestPrepare(t.GetMethodName(), false, false);
        //        appointments checkValueA = null;
        //        appointments checkValueB = new appointments();

        //        //Act
        //        checkValueB = sqlConOut.AppointmentsFind(null);

        //        //Assert
        //        TestTeardown();
        //        Assert.Equal(checkValueA, checkValueB);
        //    }
        //}
        #endregion

        #region private
        private List<string> WaitForAvailableDB()
        {
            try
            {
                for (int i = 0; i < 100; i++)
                {
                    List<string> lstMUId = sqlConSdk.UnitTest_FindAllActiveCases();

                    if (lstMUId.Count == 1)
                    {
                        return lstMUId;
                    }
                    else
                    {
                        Thread.Sleep(100);
                    }
                }
                throw new Exception("WaitForAvailableDB failed. Due to failed 100 attempts");
            }
            catch (Exception ex)
            {
                throw new Exception("WaitForAvailableDB failed", ex);
            }
        }

        private bool WaitForAvailableMicroting(int interactionCaseId)
        {
            try
            {
                string lastReply = "";

                for (int i = 0; i < 125; i++)
                {
                    var lst = sqlConSdk.UnitTest_FindAllActiveInteractionCaseLists(interactionCaseId);
                    var cas = sqlConSdk.UnitTest_FindInteractionCase(interactionCaseId);

                    if (cas.workflow_state == "failed to sync")
                        return true;

                    int missingCount = 0;

                    foreach (var item in lst)
                    {
                        if (string.IsNullOrEmpty(item.microting_uid))
                            missingCount++;
                    }

                    if (missingCount == 0)
                    {
                        lastReply = "";

                        foreach (var item in lst)
                        {
                            string reply = coreSdk.CaseCheck(item.microting_uid);

                            if (!reply.Contains("success"))
                                missingCount++;

                            lastReply += reply + " // ";
                        }

                        if (missingCount == 0)
                            return true;
                    }

                    Thread.Sleep(250 + 12 * i);
                }
                coreSdk.log.LogCritical("Not Specified", "TraceMsg:'" + lastReply.Trim() + "'");
                throw new Exception("WaitForAvailableMicroting failed. Due to failed 125 attempts (1+ min)");
            }
            catch (Exception ex)
            {
                throw new Exception("WaitForAvailableMicroting failed", ex);
            }
        }

        private string ClearXml(string inputXmlString)
        {
            inputXmlString = t.LocateReplaceAll(inputXmlString, "<StartDate>", "</StartDate>", "xxx");
            inputXmlString = t.LocateReplaceAll(inputXmlString, "<EndDate>", "</EndDate>", "xxx");
            inputXmlString = t.LocateReplaceAll(inputXmlString, "<Language>", "</Language>", "xxx");
            inputXmlString = t.LocateReplaceAll(inputXmlString, "<Id>", "</Id>", "xxx");
            inputXmlString = t.LocateReplaceAll(inputXmlString, "<DefaultValue>", "</DefaultValue>", "xxx");
            inputXmlString = t.LocateReplaceAll(inputXmlString, "<MaxValue>", "</MaxValue>", "xxx");
            inputXmlString = t.LocateReplaceAll(inputXmlString, "<MinValue>", "</MinValue>", "xxx");

            return inputXmlString;
        }

        private void CaseComplet(string microtingUId, string checkUId)
        {
            sqlConSdk.NotificationCreate(DateTime.Now.ToLongTimeString(), microtingUId, "unit_fetch");

            while (sqlConSdk.UnitTest_FindAllActiveNotifications().Count > 0)
                Thread.Sleep(100);

            sqlConSdk.NotificationCreate(DateTime.Now.ToLongTimeString(), microtingUId, "check_status");

            while (sqlConSdk.UnitTest_FindAllActiveNotifications().Count > 0)
                Thread.Sleep(100);

            if (checkUId != null)
                sqlConSdk.CaseCreate(2, siteId1, microtingUId, checkUId, "", "", DateTime.Now);

            core_UT.CaseComplet(microtingUId, checkUId);
        }

        private void InteractionCaseComplet(int interactionCaseId)
        {
            var lst = sqlConSdk.UnitTest_FindAllActiveInteractionCaseLists(interactionCaseId);

            foreach (var item in lst)
            {
                CaseComplet(item.microting_uid, null);
            }
        }

        private string LoadFil(string path)
        {
            try
            {
                lock (_lockFil)
                {
                    string str = "";
                    using (StreamReader sr = new StreamReader(path, true))
                    {
                        str = sr.ReadToEnd();
                    }
                    return str;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to load fil", ex);
            }
        }
        #endregion

        #region events
        public void EventCaseCreated(object sender, EventArgs args)
        {
            ////DOSOMETHING: changed to fit your wishes and needs 
            //Case_Dto temp = (Case_Dto)sender;
            //int siteId = temp.SiteId;
            //string caseType = temp.CaseType;
            //string caseUid = temp.CaseUId;
            //string mUId = temp.MicrotingUId;
            //string checkUId = temp.CheckUId;
        }

        public void EventCaseRetrived(object sender, EventArgs args)
        {
            ////DOSOMETHING: changed to fit your wishes and needs 
            //Case_Dto temp = (Case_Dto)sender;
            //int siteId = temp.SiteId;
            //string caseType = temp.CaseType;
            //string caseUid = temp.CaseUId;
            //string mUId = temp.MicrotingUId;
            //string checkUId = temp.CheckUId;
        }

        public void EventCaseCompleted(object sender, EventArgs args)
        {
            ////DOSOMETHING: changed to fit your wishes and needs
            //Case_Dto temp = (Case_Dto)sender;
            //int siteId = temp.SiteId;
            //string caseType = temp.CaseType;
            //string caseUid = temp.CaseUId;
            //string mUId = temp.MicrotingUId;
            //string checkUId = temp.CheckUId;
        }

        public void EventCaseDeleted(object sender, EventArgs args)
        {
            ////DOSOMETHING: changed to fit your wishes and needs
            //Case_Dto temp = (Case_Dto)sender;
            //int siteId = temp.SiteId;
            //string caseType = temp.CaseType;
            //string caseUid = temp.CaseUId;
            //string mUId = temp.MicrotingUId;
            //string checkUId = temp.CheckUId;
        }

        public void EventFileDownloaded(object sender, EventArgs args)
        {
            ////DOSOMETHING: changed to fit your wishes and needs 
            //File_Dto temp = (File_Dto)sender;
            //int siteId = temp.SiteId;
            //string caseType = temp.CaseType;
            //string caseUid = temp.CaseUId;
            //string mUId = temp.MicrotingUId;
            //string checkUId = temp.CheckUId;
            //string fileLocation = temp.FileLocation;
        }

        public void EventSiteActivated(object sender, EventArgs args)
        {
            ////DOSOMETHING: changed to fit your wishes and needs 
            //int siteId = int.Parse(sender.ToString());
        }

        public void EventNotificationNotFound(object sender, EventArgs args)
        {

        }

        public void EventException(object sender, EventArgs args)
        {
            lock (_lockFil)
            {
                File.AppendAllText(@"log\\exception.txt", sender + Environment.NewLine);
            }

            throw (Exception)sender;
        }
        #endregion
    }

    #region dummy class
    [CollectionDefinition("Database collection")]
    public class DatabaseCollection : ICollectionFixture<TestContext>
    {
        // This class has no code, and is never created. Its purpose is simply
        // to be the place to apply [CollectionDefinition] and all the
        // ICollectionFixture<> interfaces.
    }
    #endregion
}