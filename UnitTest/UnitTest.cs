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
    public class UnitTest
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
        int workerMUId = 666;
        int unitMUId = 345678;

        string connectionStringOut = "";
        string connectionStringSdk = "";
        #endregion

        #region con
        public UnitTest(TestContext testContext)
        {
            connectionStringOut = testContext.GetConnectionStringOutlook();
            connectionStringSdk     = testContext.GetConnectionStringSdk();
        }
        #endregion

        #region prepare and teardown     
        private void        TestPrepare(string testName, bool startSdk, bool startOut)
        {
            adminTool = new eFormCore.AdminTools(connectionStringSdk);
            string temp = adminTool.DbClear();
            if (temp != "")
                throw new Exception("CleanUp failed (SDK)");

            sqlConSdk = new eFormSqlController.SqlController(connectionStringSdk);
            sqlConOut = new                    SqlController(connectionStringOut);

            if (!sqlConOut.UnitTest_OutlookDatabaseClear())
                throw new Exception("CleanUp failed (Outlook)");

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

        private void        TestTeardown()
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
        public void         Test000_Basics_1a_MustAlwaysPass()
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

        [Fact]
        public void         Test000_Basics_2a_PrepareAndTeardownTestdata_True_True()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), true, true);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                checkValueB = true;

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test000_Basics_2b_PrepareAndTeardownTestdata_True_False()
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

        [Fact]
        public void         Test000_Basics_2c_PrepareAndTeardownTestdata_False_True()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, true);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                checkValueB = true;

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test000_Basics_2d_PrepareAndTeardownTestdata_False_False()
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
        [Fact]
        public void         Test001_Core_1a_Start_WithNullExpection()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "serverConnectionString is not allowed to be null or empty";
                string checkValueB = "";
                Core core = new Core();

                //Act
                try
                {
                    checkValueB = core.Start(null) + "";
                }
                catch (Exception ex)
                {
                    checkValueB = ex.InnerException.Message;
                }

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test001_Core_1b_Start_WithBlankExpection()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "serverConnectionString is not allowed to be null or empty";
                string checkValueB = "";
                Core core = new Core();

                //Act
                try
                {
                    checkValueB = core.Start("").ToString();
                }
                catch (Exception ex)
                {
                    checkValueB = ex.InnerException.Message;
                }

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test001_Core_3a_Start()
        {
            //Arrange
            TestPrepare(t.GetMethodName(), false, false);
            string checkValueA = "True";
            string checkValueB = "";
            Core core = new Core();

            //Act
            try
            {
                checkValueB = core.Start(connectionStringOut).ToString();
            }
            catch (Exception ex)
            {
                checkValueB = t.PrintException(t.GetMethodName() + " failed", ex);
            }

            //Assert
            TestTeardown();
            Assert.Equal(checkValueA, checkValueB);
        }

        [Fact]
        public void         Test001_Core_4a_IsRunning()
        {
            //Arrange
            TestPrepare(t.GetMethodName(), false, false);
            string checkValueA = "FalseTrue";
            string checkValueB = "";
            Core core = new Core();

            //Act
            try
            {
                checkValueB = core.Running().ToString();
                core.Start(connectionStringOut);
                checkValueB += core.Running().ToString();
            }
            catch (Exception ex)
            {
                checkValueB = t.PrintException(t.GetMethodName() + " failed", ex);
            }

            //Assert
            TestTeardown();
            Assert.Equal(checkValueA, checkValueB);
        }

        [Fact]
        public void         Test001_Core_5a_Close()
        {
            //Arrange
            TestPrepare(t.GetMethodName(), false, false);
            string checkValueA = "True";
            string checkValueB = "";
            Core core = new Core();

            //Act
            try
            {
                checkValueB = core.Close().ToString();
            }
            catch (Exception ex)
            {
                checkValueB = t.PrintException(t.GetMethodName() + " failed", ex);
            }

            //Assert
            TestTeardown();
            Assert.Equal(checkValueA, checkValueB);
        }

        [Fact]
        public void         Test001_Core_6a_RunningForWhileThenClose()
        {
            //Arrange
            TestPrepare(t.GetMethodName(), false, false);
            string checkValueA = "FalseTrueTrue";
            string checkValueB = "";
            Core core = new Core();

            //Act
            try
            {
                checkValueB = core.Running().ToString();
                core.Start(connectionStringOut);
                Thread.Sleep(30000);
                checkValueB += core.Running().ToString();
                Thread.Sleep(05000);
                checkValueB += core.Close().ToString();
            }
            catch (Exception ex)
            {
                checkValueB = t.PrintException(t.GetMethodName() + " failed", ex);
            }

            //Assert
            TestTeardown();
            Assert.Equal(checkValueA, checkValueB);
        }
        #endregion

        #region - test 002x - sqlController (Appointments)
        [Fact]
        public void         Test002_SqlController_1a_AppointmentsCreate_WithNullExpection()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                appointments checkValueA = null;
                appointments checkValueB = new appointments();

                //Act
                checkValueB = sqlConOut.AppointmentsFind(null);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test002_SqlController_1b_AppointmentsCreate()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                int checkValueA = 1;
                int checkValueB = -1;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlConOut.LookupRead);
                checkValueB = sqlConOut.AppointmentsCreate(appoBase);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test002_SqlController_1c_AppointmentsCreateDouble()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                int checkValueA = 2;
                int checkValueB = -1;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlConOut.LookupRead);
                checkValueB = sqlConOut.AppointmentsCreate(appoBase);
                checkValueB = sqlConOut.AppointmentsCreate(appoBase);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test002_SqlController_2a_AppointmentsCancel_WithNullExpection()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = false;
                bool checkValueB = true;

                //Act
                checkValueB = sqlConOut.AppointmentsCancel(null);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test002_SqlController_2b_AppointmentsCancel_NoMatch()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = false;
                bool checkValueB = true;

                //Act
                sqlConOut.AppointmentsCreate(new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlConOut.LookupRead));
                checkValueB = sqlConOut.AppointmentsCancel("no match");

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test002_SqlController_2c_AppointmentsCancel()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                var temp = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlConOut.LookupRead);
                sqlConOut.AppointmentsCreate(temp);
                checkValueB = sqlConOut.AppointmentsCancel("globalId");

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test002_SqlController_3a_AppointmentsFind_WithNullExpection()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                appointments checkValueA = null;
                appointments checkValueB = new appointments();

                //Act
                checkValueB = sqlConOut.AppointmentsFind(null);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test002_SqlController_3b_AppointmentsFind()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "Planned Test";
                string checkValueB = "Not the right reply";

                //Act
                sqlConOut.AppointmentsCreate(new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlConOut.LookupRead));
                var match = sqlConOut.AppointmentsFind("globalId");
                checkValueB = match.location + " " + match.subject;

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test002_SqlController_4a_AppointmentsFindOne_UnableToFind()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "";
                string checkValueB = "Not the right reply";

                //Act
                checkValueB = AppointmentsFindAll();

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test002_SqlController_4b_AppointmentsFindOne_Created()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "1Processed";
                string checkValueB = "Not the right reply";

                //Act
                sqlConOut.AppointmentsCreate(new Appointment("globalId1", DateTime.Now, 30, "Test", "Planned", "body1", false, false, sqlConOut.LookupRead));
                sqlConOut.AppointmentsCreate(new Appointment("globalId2", DateTime.Now, 30, "Test", "Planned", "body2", false, false, sqlConOut.LookupRead));
                sqlConOut.AppointmentsCreate(new Appointment("globalId3", DateTime.Now, 30, "Test", "Planned", "body3", false, false, sqlConOut.LookupRead));
                checkValueB = AppointmentsFindAll();

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test002_SqlController_4c_AppointmentsFindOne_Updated()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "0CompletedCreated";
                string checkValueB = "Not the right reply";

                //Act
                sqlConOut.AppointmentsCreate(new Appointment("globalId1", DateTime.Now, 30, "Test", "Planned", "body1", false, false, sqlConOut.LookupRead));
                sqlConOut.AppointmentsCreate(new Appointment("globalId2", DateTime.Now, 30, "Test", "Planned", "body2", false, false, sqlConOut.LookupRead));
                sqlConOut.AppointmentsCreate(new Appointment("globalId3", DateTime.Now, 30, "Test", "Planned", "body3", false, false, sqlConOut.LookupRead));

                sqlConOut.AppointmentsUpdate("globalId1", WorkflowState.Created, null, "", "");
                sqlConOut.AppointmentsUpdate("globalId2", WorkflowState.Created, null, "", "");
                sqlConOut.AppointmentsUpdate("globalId3", WorkflowState.Created, null, "", "");

                sqlConOut.AppointmentsUpdate("globalId3", WorkflowState.Completed, null, "", "");

                checkValueB = AppointmentsFindAll();

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test002_SqlController_4d_AppointmentsFindOne_Reflected()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA1 = "01Created";
                string checkValueA2 = "012CanceledRetrivedSent";
                string checkValueB1 = "Not the right reply";
                string checkValueB2 = "Not the right reply";

                //Act
                sqlConOut.AppointmentsCreate(new Appointment("globalId1", DateTime.Now, 30, "Test", "Planned", "body1", false, false, sqlConOut.LookupRead));
                sqlConOut.AppointmentsCreate(new Appointment("globalId2", DateTime.Now, 30, "Test", "Planned", "body2", false, false, sqlConOut.LookupRead));
                sqlConOut.AppointmentsCreate(new Appointment("globalId3", DateTime.Now, 30, "Test", "Planned", "body3", false, false, sqlConOut.LookupRead));

                sqlConOut.AppointmentsUpdate("globalId1", WorkflowState.Created, null, "", "");
                sqlConOut.AppointmentsUpdate("globalId2", WorkflowState.Created, null, "", "");
                sqlConOut.AppointmentsUpdate("globalId3", WorkflowState.Created, null, "", "");

                sqlConOut.AppointmentsReflected("globalId1");
                sqlConOut.AppointmentsReflected("globalId3");

                checkValueB1 = AppointmentsFindAll();

                sqlConOut.AppointmentsUpdate("globalId1", WorkflowState.Sent, null, "", "");
                sqlConOut.AppointmentsUpdate("globalId2", WorkflowState.Retrived, null, "", "");
                sqlConOut.AppointmentsUpdate("globalId3", WorkflowState.Canceled, null, "", "");

                sqlConOut.AppointmentsReflected("globalId1");
                sqlConOut.AppointmentsReflected("globalId3");
                sqlConOut.AppointmentsReflected("globalId3");
                sqlConOut.AppointmentsReflected("globalId3");
                sqlConOut.AppointmentsReflected("globalId3");

                checkValueB2 = AppointmentsFindAll();

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA1, checkValueB1);
                Assert.Equal(checkValueA2, checkValueB2);
            }
        }
        #endregion

        #region - test 003x - sqlController (Lookup)
        [Fact]
        public void         Test003_SqlController_1a_LookupCreate_Withxpection()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = false;
                bool checkValueB1 = true;
                bool checkValueB2 = true;
                bool checkValueB3 = true;
                bool checkValueB4 = true;

                //Act
                checkValueB1 = sqlConOut.LookupCreateUpdate(null, null);
                checkValueB2 = sqlConOut.LookupCreateUpdate("", null);
                checkValueB3 = sqlConOut.LookupCreateUpdate(null, "");
                checkValueB4 = sqlConOut.LookupCreateUpdate("", "");

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB1);
                Assert.Equal(checkValueA, checkValueB2);
                Assert.Equal(checkValueA, checkValueB3);
                Assert.Equal(checkValueA, checkValueB4);
            }
        }

        [Fact]
        public void         Test003_SqlController_1b_LookupCreate()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = true;
                bool checkValueB1 = false;
                bool checkValueB2 = false;
                bool checkValueB3 = false;
                bool checkValueB4 = false;

                //Act
                checkValueB1 = sqlConOut.LookupCreateUpdate("a", "1");
                checkValueB2 = sqlConOut.LookupCreateUpdate("b", "2");
                checkValueB3 = sqlConOut.LookupCreateUpdate("c", "3");
                checkValueB4 = sqlConOut.LookupCreateUpdate("d", "4");

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB1);
                Assert.Equal(checkValueA, checkValueB2);
                Assert.Equal(checkValueA, checkValueB3);
                Assert.Equal(checkValueA, checkValueB4);
            }
        }

        [Fact]
        public void         Test003_SqlController_1c_LookupCreateAndUpdate()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);

                bool checkValueA1 = true;
                bool checkValueA2 = false;

                bool checkValueB1 = false;
                bool checkValueB2 = false;
                bool checkValueB3 = false;
                bool checkValueB4 = false;
                bool checkValueB5 = false;
                bool checkValueB6 = false;
                bool checkValueB7 = false;
                bool checkValueB8 = false;
                bool checkValueB9 = false;

                //Act
                checkValueB1 = sqlConOut.LookupCreateUpdate("a", "1");
                checkValueB2 = sqlConOut.LookupCreateUpdate("b", "2");
                checkValueB3 = sqlConOut.LookupCreateUpdate("c", "3");
                checkValueB4 = sqlConOut.LookupCreateUpdate("c", "4");
                checkValueB5 = sqlConOut.LookupCreateUpdate("b", "5");
                checkValueB6 = sqlConOut.LookupCreateUpdate("d", "6");
                checkValueB7 = sqlConOut.LookupCreateUpdate("", "4");
                checkValueB8 = sqlConOut.LookupCreateUpdate("c", null);
                checkValueB9 = sqlConOut.LookupCreateUpdate("c", "9");

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA1, checkValueB1);
                Assert.Equal(checkValueA1, checkValueB2);
                Assert.Equal(checkValueA1, checkValueB3);
                Assert.Equal(checkValueA1, checkValueB4);
                Assert.Equal(checkValueA1, checkValueB5);
                Assert.Equal(checkValueA1, checkValueB6);
                Assert.Equal(checkValueA2, checkValueB7);
                Assert.Equal(checkValueA2, checkValueB8);
                Assert.Equal(checkValueA1, checkValueB9);
            }
        }

        [Fact]
        public void         Test003_SqlController_2a_LookupRead()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);

                string checkValueA1 = "1A1b1k";
                string checkValueA2 = "2C2d2m";

                string checkValueB1 = "";
                string checkValueB2 = "";

                //Act
                sqlConOut.LookupCreateUpdate("Ab", "1A1b1k");
                sqlConOut.LookupCreateUpdate("CD", "2C2d2m");

                checkValueB1 = sqlConOut.LookupRead("aB");
                checkValueB2 = sqlConOut.LookupRead("Cd");

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA1, checkValueB1);
                Assert.Equal(checkValueA2, checkValueB2);
                Assert.NotEqual(checkValueA1, checkValueB2.ToUpper());
                Assert.NotEqual(checkValueA1, checkValueB2.ToLower());
                Assert.NotEqual(checkValueA2, checkValueB2.ToUpper());
                Assert.NotEqual(checkValueA2, checkValueB2.ToLower());
            }
        }

        [Fact]
        public void         Test003_SqlController_3a_LookupReadAll()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "1A1b1k2C2d2m3e3F3k4G4H4m";
                string checkValueB = "";

                //Act
                sqlConOut.LookupCreateUpdate("Ab", "1A1b1k");
                sqlConOut.LookupCreateUpdate("CD", "2C2d2m");
                sqlConOut.LookupCreateUpdate("EF", "3e3F3k");
                sqlConOut.LookupCreateUpdate("GH", "4G4H4m");

                var lst = sqlConOut.LookupReadAll();

                foreach (var item in lst)
                    checkValueB += item.value;

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test003_SqlController_4a_LookupDelete()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "1A1b1k3e3F3k1A1b1k";
                string checkValueB = "";

                //Act
                sqlConOut.LookupCreateUpdate("Ab", "1A1b1k");
                sqlConOut.LookupCreateUpdate("CD", "2C2d2m");
                sqlConOut.LookupCreateUpdate("EF", "3e3F3k");
                sqlConOut.LookupCreateUpdate("GH", "4G4H4m");

                sqlConOut.LookupDelete("Cd");
                sqlConOut.LookupDelete("gH");

                var lst = sqlConOut.LookupReadAll();

                foreach (var item in lst)
                    checkValueB += item.value;

                sqlConOut.LookupDelete("ef");

                lst = sqlConOut.LookupReadAll();

                foreach (var item in lst)
                    checkValueB += item.value;

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }
        #endregion

        #region - test 004x - sqlController (SDK)
        [Fact]
        public void         Test004_SqlController_1a_SyncInteractionCase()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                int checkValueA = 1;
                int checkValueB = 1;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlConOut.LookupRead);
                sqlConOut.AppointmentsCreate(appoBase);
                sqlConOut.SyncInteractionCase();

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test004_SqlController_2a_InteractionCaseCreate()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlConOut.LookupRead);
                int id = sqlConOut.AppointmentsCreate(appoBase);
                var app = sqlConOut.AppointmentsFind("globalId");

                checkValueB = sqlConOut.InteractionCaseCreate(app);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test004_SqlController_3a_InteractionCaseDelete()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlConOut.LookupRead);
                int id = sqlConOut.AppointmentsCreate(appoBase);
                var app = sqlConOut.AppointmentsFind("globalId");

                checkValueB = sqlConOut.InteractionCaseCreate(app);
                //checkValueB = sqlConOut.InteractionCaseDelete(app); Lacks to fake a SDK sending, so it can be delete. Needs to make more test, for deletions for different stages

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test004_SqlController_4a_InteractionCaseDelete()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlConOut.LookupRead);
                int id = sqlConOut.AppointmentsCreate(appoBase);
                var app = sqlConOut.AppointmentsFind("globalId");

                checkValueB = sqlConOut.InteractionCaseCreate(app);
                //checkValueB = sqlConOut.InteractionCaseDelete(app); Lacks to fake a SDK sending, so it can be delete. Needs to make more test, for deletions for different stages

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void         Test004_SqlController_5a_InteractionCaseProcessed_NotMade()
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

        [Fact]
        public void         Test004_SqlController_6a_SiteLookupName_NotMade()
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

        #region - test 005x - sqlController (Settings)
        //Not active, as would fuck up the stat of settings
        //[Fact]
        //public void         Test005_SqlController_1a_SettingCreateDefaults()
        //{
        //    lock (_lockTest)
        //    {
        //        //Arrange
        //        TestPrepare(t.GetMethodName(), false, false);
        //        bool checkValueA = true;
        //        bool checkValueB = false;

        //        //Act
        //        checkValueB = sqlConOut.SettingCreateDefaults();

        //        //Assert
        //        TestTeardown();
        //        Assert.Equal(checkValueA, checkValueB);
        //    }
        //}

        //[Fact]
        //public void         Test005_SqlController_2a_SettingCreate()
        //{
        //    lock (_lockTest)
        //    {
        //        //Arrange
        //        TestPrepare(t.GetMethodName(), false, false);
        //        bool checkValueA1 = true;
        //        bool checkValueA2 = true;
        //        bool checkValueB1 = false;
        //        bool checkValueB2 = false;

        //        //Act
        //        checkValueB1 = sqlConOut.SettingCreate(Settings.firstRunDone);
        //        checkValueB2 = sqlConOut.SettingCreate(Settings.logLevel);

        //        //Assert
        //        TestTeardown();
        //        Assert.Equal(checkValueA1, checkValueB1);
        //        Assert.Equal(checkValueA2, checkValueB2);
        //    }
        //}

        //Not active, as would fuck up the stat of settings
        [Fact]
        public void         Test005_SqlController_3a_SettingRead()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA1 = "true";
                string checkValueA2 = "4";
                string checkValueB1 = "";
                string checkValueB2 = "";

                //Act
                sqlConOut.SettingCreate(Settings.firstRunDone);
                sqlConOut.SettingCreate(Settings.logLevel);

                checkValueB1 = sqlConOut.SettingRead(Settings.firstRunDone);
                checkValueB2 = sqlConOut.SettingRead(Settings.logLevel);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA1, checkValueB1);
                Assert.Equal(checkValueA2, checkValueB2);
            }
        }

        //Not active, as would fuck up the stat of settings
        //[Fact]
        //public void         Test005_SqlController_4a_SettingUpdate()
        //{       
        //    lock (_lockTest)
        //    {
        //        //Arrange
        //        TestPrepare(t.GetMethodName(), false, false);
        //        string checkValueA = "tempValuefinalValue";
        //        string checkValueB1 = "";
        //        string checkValueB2 = "";

        //        //Act
        //        sqlConOut.SettingCreate(Settings.firstRunDone);
        //        sqlConOut.SettingCreate(Settings.logLevel);

        //        sqlConOut.SettingUpdate(Settings.firstRunDone, "tempValue");
        //        sqlConOut.SettingUpdate(Settings.logLevel, "tempValue");

        //        checkValueB1 = sqlConOut.SettingRead(Settings.firstRunDone);
        //        checkValueB2 = sqlConOut.SettingRead(Settings.logLevel);

        //        sqlConOut.SettingUpdate(Settings.firstRunDone, "finalValue");
        //        sqlConOut.SettingUpdate(Settings.logLevel, "finalValue");

        //        checkValueB1 += sqlConOut.SettingRead(Settings.firstRunDone);
        //        checkValueB2 += sqlConOut.SettingRead(Settings.logLevel);

        //        //Assert
        //        TestTeardown();
        //        Assert.Equal(checkValueA, checkValueB1);
        //        Assert.Equal(checkValueA, checkValueB2);
        //    }
        //}

        [Fact]
        public void         Test005_SqlController_5a_SettingCheckAll()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                int checkValueA = 0;
                int checkValueB = -1;

                //Act
                sqlConOut.SettingCreateDefaults();
                var temp = sqlConOut.SettingCheckAll();
                checkValueB = temp.Count();

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }
        #endregion

        #region private
        private string      AppointmentsFindAll()
        {
            string returnValue = "";

            if (sqlConOut.AppointmentsFindOne(0) != null)                                   returnValue += "0";
            if (sqlConOut.AppointmentsFindOne(1) != null)                                   returnValue += "1";
            if (sqlConOut.AppointmentsFindOne(2) != null)                                   returnValue += "2";
            if (sqlConOut.AppointmentsFindOne(3) != null)                                   returnValue += "3";
            if (sqlConOut.AppointmentsFindOne(4) != null)                                   returnValue += "4";

            if (sqlConOut.AppointmentsFindOne(WorkflowState.Canceled) != null)              returnValue += "Canceled";
            if (sqlConOut.AppointmentsFindOne(WorkflowState.Completed) != null)             returnValue += "Completed";
            if (sqlConOut.AppointmentsFindOne(WorkflowState.Created) != null)               returnValue += "Created";
            if (sqlConOut.AppointmentsFindOne(WorkflowState.Failed_to_expection) != null)   returnValue += "Failed_to_expection";
            if (sqlConOut.AppointmentsFindOne(WorkflowState.Failed_to_intrepid) != null)    returnValue += "Failed_to_intrepid";
            if (sqlConOut.AppointmentsFindOne(WorkflowState.Planned) != null)               returnValue += "Planned";
            if (sqlConOut.AppointmentsFindOne(WorkflowState.Processed) != null)             returnValue += "Processed";
            if (sqlConOut.AppointmentsFindOne(WorkflowState.Retrived) != null)              returnValue += "Retrived";
            if (sqlConOut.AppointmentsFindOne(WorkflowState.Revoked) != null)               returnValue += "Revoked";
            if (sqlConOut.AppointmentsFindOne(WorkflowState.Sent) != null)                  returnValue += "Sent";

            return returnValue;
        }

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

        private bool        WaitForAvailableMicroting(int interactionCaseId)
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

        private string      ClearXml(string inputXmlString)
        {
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<StartDate>", "</StartDate>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<EndDate>", "</EndDate>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<Language>", "</Language>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<Id>", "</Id>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<DefaultValue>", "</DefaultValue>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<MaxValue>", "</MaxValue>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<MinValue>", "</MinValue>", "xxx", true);

            return inputXmlString;
        }

        private void        CaseComplet(string microtingUId, string checkUId)
        {
            sqlConSdk.NotificationCreate(DateTime.Now.ToLongTimeString(), microtingUId, "unit_fetch");

            while (sqlConSdk.UnitTest_FindAllActiveNotifications().Count > 0)
                Thread.Sleep(100);

            sqlConSdk.NotificationCreate(DateTime.Now.ToLongTimeString(), microtingUId, "check_status");

            while (sqlConSdk.UnitTest_FindAllActiveNotifications().Count > 0)
                Thread.Sleep(100);

            if (checkUId != null)
                sqlConSdk.CaseCreate(2, siteId1, microtingUId, checkUId, "", "", DateTime.Now);

            core_UT.CaseComplet(microtingUId, checkUId, workerMUId, unitMUId);
        }

        private void        InteractionCaseComplet(int interactionCaseId)
        {
            var lst = sqlConSdk.UnitTest_FindAllActiveInteractionCaseLists(interactionCaseId);

            foreach (var item in lst)
            {
                CaseComplet(item.microting_uid, null);
            }
        }

        private string      LoadFil(string path)
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