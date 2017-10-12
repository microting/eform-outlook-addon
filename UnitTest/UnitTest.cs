using eFormShared;
using OutlookSql;
using OutlookCore;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Xunit;
using System.Threading;
using OutlookOffice;

namespace UnitTest
{
    public class TestContext : IDisposable
    {
        //

        string connectionStringLocal_SDK = "Data Source=DESKTOP-7V1APE5\\SQLEXPRESS;Initial Catalog=" + "UnitTest_Outlook_" + "Microting"        + ";Integrated Security=True";
        string connectionStringLocal_OUT = "Data Source=DESKTOP-7V1APE5\\SQLEXPRESS;Initial Catalog=" + "UnitTest_Outlook_" + "MicrotingOutlook" + ";Integrated Security=True";

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
                if (Environment.MachineName.ToLower().Contains("testing") || Environment.MachineName.ToLower().Contains("travis"))
                {
                    serverConnectionString_SDK = "Persist Security Info=True;server=localhost;database=" + "UnitTest_Outlook_" + "Microting"        + ";uid=root;password="; //Uses travis database
                    serverConnectionString_OUT = "Persist Security Info=True;server=localhost;database=" + "UnitTest_Outlook_" + "MicrotingOutlook" + ";uid=root;password="; //Uses travis database
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
        Core coreOut;
        CoreUnitTest coreOut_UT;
        eFormCore.Core coreSdk;
        eFormCore.CoreUnitTest coreSdk_UT;
        eFormSqlController.SqlController sqlConSdk;
        SqlController sqlController;
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
            connectionStringSdk = testContext.GetConnectionStringSdk();
        }
        #endregion

        #region prepare and teardown     
        private void TestPrepare(string testName, bool startOut, bool startSdk)
        {
            adminTool = new eFormCore.AdminTools(connectionStringSdk);
            string temp = adminTool.DbClear();
            if (temp != "")
                throw new Exception("CleanUp failed (SDK)");


            sqlConSdk = new eFormSqlController.SqlController(connectionStringSdk);
            sqlController = new SqlController(connectionStringOut);
            sqlConSdk.UnitTest_TruncateTable(nameof(logs));
            sqlConSdk.UnitTest_TruncateTable(nameof(log_exceptions));
            sqlController.UnitTest_TruncateTable(nameof(logs));
            sqlController.UnitTest_TruncateTable(nameof(log_exceptions));

            if (!sqlController.UnitTest_OutlookDatabaseClear())
                throw new Exception("CleanUp failed (Outlook)");

            coreSdk = new eFormCore.Core();
            coreSdk_UT = new eFormCore.CoreUnitTest(coreSdk);
            coreOut = new Core();
            coreOut_UT = new CoreUnitTest(coreOut);

            coreOut.HandleEventException += EventException;

            if (startSdk)
                coreSdk.Start(connectionStringSdk);

            if (startOut)
                coreOut.Start(connectionStringOut);
        }

        private void TestTeardown()
        {
            if (coreSdk != null)
                if (coreSdk.Running())
                    coreSdk_UT.Close();

            if (coreOut != null)
                if (coreOut.Running())
                    coreOut_UT.Close();
        }
        #endregion

        #region - test 000x - virtal basics
        [Fact]
        public void Test000_Basics_0a_EnvironmentMachineName()
        {
            lock (_lockTest)
            {
                //Arrange
                string checkValueA = "";
                string checkValueB = Environment.MachineName;

                //Act

                //Assert
                Assert.Equal(checkValueA, checkValueB);
            }
        }

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

        [Fact]
        public void Test000_Basics_2a_PrepareAndTeardownTestdata_True_True()
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
        public void Test000_Basics_2b_PrepareAndTeardownTestdata_True_False()
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
        public void Test000_Basics_2c_PrepareAndTeardownTestdata_False_True()
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
        public void Test000_Basics_2d_PrepareAndTeardownTestdata_False_False()
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

        #region - test 001x - core
        [Fact]
        public void Test001_Core_1a_Start_WithNullExpection()
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
        public void Test001_Core_1b_Start_WithBlankExpection()
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
        public void Test001_Core_3a_Start()
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
        public void Test001_Core_4a_IsRunning()
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
        public void Test001_Core_5a_Close()
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
        public void Test001_Core_6a_RunningForWhileThenClose()
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
        public void Test002_SqlController_1a_AppointmentsCreate_WithNullExpection()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                appointments checkValueA = null;
                appointments checkValueB = new appointments();

                //Act
                checkValueB = sqlController.AppointmentsFind(null);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test002_SqlController_1b_AppointmentsCreate()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                int checkValueA = 1;
                int checkValueB = -1;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlController.LookupRead);
                checkValueB = sqlController.AppointmentsCreate(appoBase);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test002_SqlController_1c_AppointmentsCreateDouble()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                int checkValueA = 2;
                int checkValueB = -1;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlController.LookupRead);
                checkValueB = sqlController.AppointmentsCreate(appoBase);
                checkValueB = sqlController.AppointmentsCreate(appoBase);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test002_SqlController_2a_AppointmentsCancel_WithNullExpection()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = false;
                bool checkValueB = true;

                //Act
                checkValueB = sqlController.AppointmentsCancel(null);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test002_SqlController_2b_AppointmentsCancel_NoMatch()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = false;
                bool checkValueB = true;

                //Act
                sqlController.AppointmentsCreate(new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlController.LookupRead));
                checkValueB = sqlController.AppointmentsCancel("no match");

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test002_SqlController_2c_AppointmentsCancel()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                var temp = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlController.LookupRead);
                sqlController.AppointmentsCreate(temp);
                checkValueB = sqlController.AppointmentsCancel("globalId");

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test002_SqlController_3a_AppointmentsFind_WithNullExpection()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                appointments checkValueA = null;
                appointments checkValueB = new appointments();

                //Act
                checkValueB = sqlController.AppointmentsFind(null);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test002_SqlController_3b_AppointmentsFind()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "Planned Test";
                string checkValueB = "Not the right reply";

                //Act
                sqlController.AppointmentsCreate(new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlController.LookupRead));
                var match = sqlController.AppointmentsFind("globalId");
                checkValueB = match.location + " " + match.subject;

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test002_SqlController_4a_AppointmentsFindOne_UnableToFind()
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
        public void Test002_SqlController_4b_AppointmentsFindOne_Created()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "1Processed";
                string checkValueB = "Not the right reply";

                //Act
                sqlController.AppointmentsCreate(new Appointment("globalId1", DateTime.Now, 30, "Test", "Planned", "body1", false, false, sqlController.LookupRead));
                sqlController.AppointmentsCreate(new Appointment("globalId2", DateTime.Now, 30, "Test", "Planned", "body2", false, false, sqlController.LookupRead));
                sqlController.AppointmentsCreate(new Appointment("globalId3", DateTime.Now, 30, "Test", "Planned", "body3", false, false, sqlController.LookupRead));
                checkValueB = AppointmentsFindAll();

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test002_SqlController_4c_AppointmentsFindOne_Updated()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "0CompletedCreated";
                string checkValueB = "Not the right reply";

                //Act
                sqlController.AppointmentsCreate(new Appointment("globalId1", DateTime.Now, 30, "Test", "Planned", "body1", false, false, sqlController.LookupRead));
                sqlController.AppointmentsCreate(new Appointment("globalId2", DateTime.Now, 30, "Test", "Planned", "body2", false, false, sqlController.LookupRead));
                sqlController.AppointmentsCreate(new Appointment("globalId3", DateTime.Now, 30, "Test", "Planned", "body3", false, false, sqlController.LookupRead));

                sqlController.AppointmentsUpdate("globalId1", WorkflowState.Created, null, "", "");
                sqlController.AppointmentsUpdate("globalId2", WorkflowState.Created, null, "", "");
                sqlController.AppointmentsUpdate("globalId3", WorkflowState.Created, null, "", "");

                sqlController.AppointmentsUpdate("globalId3", WorkflowState.Completed, null, "", "");

                checkValueB = AppointmentsFindAll();

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test002_SqlController_4d_AppointmentsFindOne_Reflected()
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
                sqlController.AppointmentsCreate(new Appointment("globalId1", DateTime.Now, 30, "Test", "Planned", "body1", false, false, sqlController.LookupRead));
                sqlController.AppointmentsCreate(new Appointment("globalId2", DateTime.Now, 30, "Test", "Planned", "body2", false, false, sqlController.LookupRead));
                sqlController.AppointmentsCreate(new Appointment("globalId3", DateTime.Now, 30, "Test", "Planned", "body3", false, false, sqlController.LookupRead));

                sqlController.AppointmentsUpdate("globalId1", WorkflowState.Created, null, "", "");
                sqlController.AppointmentsUpdate("globalId2", WorkflowState.Created, null, "", "");
                sqlController.AppointmentsUpdate("globalId3", WorkflowState.Created, null, "", "");

                sqlController.AppointmentsReflected("globalId1");
                sqlController.AppointmentsReflected("globalId3");

                checkValueB1 = AppointmentsFindAll();

                sqlController.AppointmentsUpdate("globalId1", WorkflowState.Sent, null, "", "");
                sqlController.AppointmentsUpdate("globalId2", WorkflowState.Retrived, null, "", "");
                sqlController.AppointmentsUpdate("globalId3", WorkflowState.Canceled, null, "", "");

                sqlController.AppointmentsReflected("globalId1");
                sqlController.AppointmentsReflected("globalId3");
                sqlController.AppointmentsReflected("globalId3");
                sqlController.AppointmentsReflected("globalId3");
                sqlController.AppointmentsReflected("globalId3");

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
        public void Test003_SqlController_1a_LookupCreate_Withxpection()
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
                checkValueB1 = sqlController.LookupCreateUpdate(null, null);
                checkValueB2 = sqlController.LookupCreateUpdate("", null);
                checkValueB3 = sqlController.LookupCreateUpdate(null, "");
                checkValueB4 = sqlController.LookupCreateUpdate("", "");

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB1);
                Assert.Equal(checkValueA, checkValueB2);
                Assert.Equal(checkValueA, checkValueB3);
                Assert.Equal(checkValueA, checkValueB4);
            }
        }

        [Fact]
        public void Test003_SqlController_1b_LookupCreate()
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
                checkValueB1 = sqlController.LookupCreateUpdate("a", "1");
                checkValueB2 = sqlController.LookupCreateUpdate("b", "2");
                checkValueB3 = sqlController.LookupCreateUpdate("c", "3");
                checkValueB4 = sqlController.LookupCreateUpdate("d", "4");

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB1);
                Assert.Equal(checkValueA, checkValueB2);
                Assert.Equal(checkValueA, checkValueB3);
                Assert.Equal(checkValueA, checkValueB4);
            }
        }

        [Fact]
        public void Test003_SqlController_1c_LookupCreateAndUpdate()
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
                checkValueB1 = sqlController.LookupCreateUpdate("a", "1");
                checkValueB2 = sqlController.LookupCreateUpdate("b", "2");
                checkValueB3 = sqlController.LookupCreateUpdate("c", "3");
                checkValueB4 = sqlController.LookupCreateUpdate("c", "4");
                checkValueB5 = sqlController.LookupCreateUpdate("b", "5");
                checkValueB6 = sqlController.LookupCreateUpdate("d", "6");
                checkValueB7 = sqlController.LookupCreateUpdate("", "4");
                checkValueB8 = sqlController.LookupCreateUpdate("c", null);
                checkValueB9 = sqlController.LookupCreateUpdate("c", "9");

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
        public void Test003_SqlController_2a_LookupRead()
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
                sqlController.LookupCreateUpdate("Ab", "1A1b1k");
                sqlController.LookupCreateUpdate("CD", "2C2d2m");

                checkValueB1 = sqlController.LookupRead("aB");
                checkValueB2 = sqlController.LookupRead("Cd");

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
        public void Test003_SqlController_3a_LookupReadAll()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "1A1b1k2C2d2m3e3F3k4G4H4m";
                string checkValueB = "";

                //Act
                sqlController.LookupCreateUpdate("Ab", "1A1b1k");
                sqlController.LookupCreateUpdate("CD", "2C2d2m");
                sqlController.LookupCreateUpdate("EF", "3e3F3k");
                sqlController.LookupCreateUpdate("GH", "4G4H4m");

                var lst = sqlController.LookupReadAll();

                foreach (var item in lst)
                    checkValueB += item.value;

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test003_SqlController_4a_LookupDelete()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "1A1b1k3e3F3k1A1b1k";
                string checkValueB = "";

                //Act
                sqlController.LookupCreateUpdate("Ab", "1A1b1k");
                sqlController.LookupCreateUpdate("CD", "2C2d2m");
                sqlController.LookupCreateUpdate("EF", "3e3F3k");
                sqlController.LookupCreateUpdate("GH", "4G4H4m");

                sqlController.LookupDelete("Cd");
                sqlController.LookupDelete("gH");

                var lst = sqlController.LookupReadAll();

                foreach (var item in lst)
                    checkValueB += item.value;

                sqlController.LookupDelete("ef");

                lst = sqlController.LookupReadAll();

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
        public void Test004_SqlController_1a_SyncInteractionCase()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                int checkValueA = 1;
                int checkValueB = 1;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlController.LookupRead);
                sqlController.AppointmentsCreate(appoBase);
                sqlController.SyncInteractionCase();

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test004_SqlController_2a_InteractionCaseCreate()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlController.LookupRead);
                int id = sqlController.AppointmentsCreate(appoBase);
                var app = sqlController.AppointmentsFind("globalId");

                checkValueB = sqlController.InteractionCaseCreate(app);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test004_SqlController_3a_InteractionCaseDelete()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlController.LookupRead);
                int id = sqlController.AppointmentsCreate(appoBase);
                var app = sqlController.AppointmentsFind("globalId");

                checkValueB = sqlController.InteractionCaseCreate(app);
                //checkValueB = sqlConOut.InteractionCaseDelete(app); Lacks to fake a SDK sending, so it can be delete. Needs to make more test, for deletions for different stages

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test004_SqlController_4a_InteractionCaseDelete()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = true;
                bool checkValueB = false;

                //Act
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlController.LookupRead);
                int id = sqlController.AppointmentsCreate(appoBase);
                var app = sqlController.AppointmentsFind("globalId");

                checkValueB = sqlController.InteractionCaseCreate(app);
                //checkValueB = sqlConOut.InteractionCaseDelete(app); Lacks to fake a SDK sending, so it can be delete. Needs to make more test, for deletions for different stages

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test004_SqlController_5a_InteractionCaseProcessed_NotMade()
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
        public void Test004_SqlController_6a_SiteLookupName_NotMade()
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
        //Not active, as would fuck up the stat of settings. Don't run unless settings stat is not improtant
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
        public void Test005_SqlController_3a_SettingRead()
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
                sqlController.SettingCreate(Settings.firstRunDone);
                sqlController.SettingCreate(Settings.logLevel);

                checkValueB1 = sqlController.SettingRead(Settings.firstRunDone);
                checkValueB2 = sqlController.SettingRead(Settings.logLevel);

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA1, checkValueB1);
                Assert.Equal(checkValueA2, checkValueB2);
            }
        }

        //Not active, as would fuck up the stat of settings. Don't run unless settings stat is not improtant
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
        public void Test005_SqlController_5a_SettingCheckAll()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                int checkValueA = 0;
                int checkValueB = -1;

                //Act
                sqlController.SettingCreateDefaults();
                var temp = sqlController.SettingCheckAll();
                checkValueB = temp.Count();

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }
        #endregion

        #region - test 006x - outlookController
        [Fact]
        public void Test006_OutlookController_1a_CalendarItemConvertRecurrences()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "true";
                string checkValueB = "";
                IOutlookController oCon = new OutlookController_Fake(sqlController, new Log(coreOut, new LogWriter(), 4));

                //Act
                bool response;
                for (int i = 0; i < 10; i++)
                    response = oCon.CalendarItemConvertRecurrences();
                checkValueB = "true";

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test006_OutlookController_2a_CalendarItemIntrepid()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                string checkValueA = "true";
                string checkValueB = "";
                IOutlookController oCon = new OutlookController_Fake(sqlController, new Log(coreOut, new LogWriter(), 4));

                //Act
                bool response;
                for (int i = 0; i < 10; i++)
                    response = oCon.CalendarItemIntrepid();
                checkValueB = "true";

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }

        [Fact]
        public void Test006_OutlookController_3a_CalendarItemReflecting()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool? checkValueA1 = null;
                bool? checkValueA2 = false;
                bool? checkValueA3 = true;
                bool? checkValueA4 = null;
                bool? checkValueA5 = true;

                bool? checkValueB1 = null;
                bool? checkValueB2 = true;
                bool? checkValueB3 = false;
                bool? checkValueB4 = true;
                bool? checkValueB5 = false;
                IOutlookController oCon = new OutlookController_Fake(sqlController, new Log(coreOut, new LogWriter(), 4));

                //Act
                checkValueB1 = oCon.CalendarItemReflecting(null);
                checkValueB2 = oCon.CalendarItemReflecting("");
                checkValueB3 = oCon.CalendarItemReflecting("pass");
                try
                {
                    checkValueB4 = oCon.CalendarItemReflecting("throw new expection");
                }
                catch
                {
                    checkValueB4 = null;
                }
                checkValueB5 = oCon.CalendarItemReflecting("other");

                //Assert
                TestTeardown();
                Assert.NotEqual(checkValueA1, checkValueB1);
                Assert.Equal(checkValueA2, checkValueB2);
                Assert.Equal(checkValueA3, checkValueB3);
                Assert.Equal(checkValueA4, checkValueB4);
                Assert.Equal(checkValueA5, checkValueB5);
            }
        }

        [Fact]
        public void Test006_OutlookController_4a_CalendarItemUpdate()
        {
            lock (_lockTest)
            {
                //Arrange
                TestPrepare(t.GetMethodName(), false, false);
                bool checkValueA = true;
                bool checkValueB = false;
                Appointment appoBase = new Appointment("globalId", DateTime.Now, 30, "Test", "Planned", "body", false, false, sqlController.LookupRead);
                IOutlookController oCon = new OutlookController_Fake(sqlController, new Log(coreOut, new LogWriter(), 4));

                //Act
                oCon.CalendarItemUpdate(appoBase, WorkflowState.Processed, false);
                oCon.CalendarItemUpdate(appoBase, WorkflowState.Processed, true);
                oCon.CalendarItemUpdate(appoBase, WorkflowState.Created, false);
                oCon.CalendarItemUpdate(appoBase, WorkflowState.Created, true);
                oCon.CalendarItemUpdate(appoBase, WorkflowState.Failed_to_expection, false);
                oCon.CalendarItemUpdate(appoBase, WorkflowState.Failed_to_intrepid, true);
                checkValueB = true;

                //Assert
                TestTeardown();
                Assert.Equal(checkValueA, checkValueB);
            }
        }
        #endregion

        #region - test 007x - core (Exception handling)
        [Fact]
        public void Test007_Core_1a_ExceptionHandling()
        {
            #region //Arrange
            TestPrepare(t.GetMethodName(), true, false);
            string checkValueA1 = "1:100000/100000/10000/0";
            string checkValueA2 = "1:010000/010000/01000/0";
            string checkValueA3 = "1:001000/001000/00100/0";
            string checkValueA4 = "1:000100/000100/00010/0";
            string checkValueB1 = "";
            string checkValueB2 = "";
            string checkValueB3 = "";
            string checkValueB4 = "";
            string tempValue = "";
            #endregion

            //Act
            try
            {
                for (int i = 0; i < 4; i++)
                {
                    coreOut.outlookController.UnitTest_ForceException("throw new Exception");
                    tempValue += WaitForRestart();
                }
            }
            catch (Exception ex)
            {
                tempValue = t.PrintException(t.GetMethodName() + " failed", ex);
            }

            #region //Assert
            TestTeardown();

            tempValue = tempValue.Replace("\r", "").Replace("\n", "");
            checkValueB1 = tempValue.Substring(0, 23);
            checkValueB2 = tempValue.Substring(23, 23);
            checkValueB3 = tempValue.Substring(46, 23);
            checkValueB4 = tempValue.Substring(69, 23);

            Assert.Equal(checkValueA1, checkValueB1);
            Assert.Equal(checkValueA2, checkValueB2);
            Assert.Equal(checkValueA3, checkValueB3);
            Assert.Equal(checkValueA4, checkValueB4);
            #endregion
        }

        [Fact]
        public void Test007_Core_2a_DoubleExceptionHandling()
        {
            #region //Arrange
            TestPrepare(t.GetMethodName(), true, false);
            string checkValueA1 = "1:100000/100000/10000/0";
            string checkValueA2 = "1:010000/010000/01000/0";
            string checkValueA3 = "1:010000/010000/01000/0";
            string checkValueA4 = "1:001000/001000/00100/0";
            string checkValueA5 = "1:001000/001000/00100/0";
            string checkValueA6 = "1:000100/000100/00010/0";
            string checkValueA7 = "1:000100/000100/00010/0";
            string checkValueB1 = "";
            string checkValueB2 = "";
            string checkValueB3 = "";
            string checkValueB4 = "";
            string checkValueB5 = "";
            string checkValueB6 = "";
            string checkValueB7 = "";
            string tempValue = "";
            #endregion

            //Act
            try
            {
                coreOut.outlookController.UnitTest_ForceException("throw new Exception");
                tempValue += WaitForRestart();

                for (int i = 0; i < 3; i++)
                {
                    coreOut.outlookController.UnitTest_ForceException("throw new Exception");
                    tempValue += WaitForRestart();

                    coreOut.outlookController.UnitTest_ForceException("throw other Exception");
                    tempValue += WaitForRestart();
                }
            }
            catch (Exception ex)
            {
                tempValue = t.PrintException(t.GetMethodName() + " failed", ex);
            }

            #region //Assert
            TestTeardown();

            tempValue = tempValue.Replace("\r", "").Replace("\n", "");
            checkValueB1 = tempValue.Substring(0, 23);
            checkValueB2 = tempValue.Substring(23, 23);
            checkValueB3 = tempValue.Substring(46, 23);
            checkValueB4 = tempValue.Substring(69, 23);
            checkValueB5 = tempValue.Substring(92, 23);
            checkValueB6 = tempValue.Substring(115, 23);
            checkValueB7 = tempValue.Substring(138, 23);
            
            Assert.Equal(checkValueA1, checkValueB1);
            Assert.Equal(checkValueA2, checkValueB2);
            Assert.Equal(checkValueA3, checkValueB3);
            Assert.Equal(checkValueA4, checkValueB4);
            Assert.Equal(checkValueA5, checkValueB5);
            Assert.Equal(checkValueA6, checkValueB6);
            Assert.Equal(checkValueA7, checkValueB7);
            #endregion
        }

        [Fact]
        public void Test007_Core_3a_FatalExceptionHandling()
        {
            #region //Arrange
            TestPrepare(t.GetMethodName(), true, false);
            string checkValueA1 = "1:100000/100000/10000/0";
            string checkValueA2 = "1:010000/010000/01000/0";
            string checkValueA3 = "1:010000/010000/01000/0";
            string checkValueA4 = "1:001000/001000/00100/0";
            string checkValueA5 = "1:001000/001000/00100/0";
            string checkValueA6 = "1:000100/000100/00010/0";
            string checkValueA7 = "1:000100/000100/00010/0";
            string checkValueA8 = "2:000000/000020/00001/1";
            string checkValueB1 = "";
            string checkValueB2 = "";
            string checkValueB3 = "";
            string checkValueB4 = "";
            string checkValueB5 = "";
            string checkValueB6 = "";
            string checkValueB7 = "";
            string checkValueB8 = "";
            string tempValue = "";
            #endregion

            //Act
            try
            {
                #region core.CaseCreate(main1, null, siteId1);
                for (int i = 0; i < 2; i++)
                {
                    coreOut.outlookController.UnitTest_ForceException("throw new Exception");
                    tempValue += WaitForRestart();
                }
                #endregion

                #region core.CaseCreate(main2, null, siteId1);
                #endregion
                #region core.CaseCreate(main1, null, siteId1);
                for (int i = 0; i < 3; i++)
                {
                    coreOut.outlookController.UnitTest_ForceException("throw other Exception");
                    tempValue += WaitForRestart();

                    coreOut.outlookController.UnitTest_ForceException("throw new Exception");
                    tempValue += WaitForRestart();
                }
                #endregion
            }
            catch (Exception ex)
            {
                tempValue = t.PrintException(t.GetMethodName() + " failed", ex);
            }

            #region //Assert
            TestTeardown();

            tempValue = tempValue.Replace("\r", "").Replace("\n", "");
            checkValueB1 = tempValue.Substring(0, 23);
            checkValueB2 = tempValue.Substring(23, 23);
            checkValueB3 = tempValue.Substring(46, 23);
            checkValueB4 = tempValue.Substring(69, 23);
            checkValueB5 = tempValue.Substring(92, 23);
            checkValueB6 = tempValue.Substring(115, 23);
            checkValueB7 = tempValue.Substring(138, 23);
            checkValueB8 = tempValue.Substring(161, 23);
            
            Assert.Equal(checkValueA1, checkValueB1);
            Assert.Equal(checkValueA2, checkValueB2);
            Assert.Equal(checkValueA3, checkValueB3);
            Assert.Equal(checkValueA4, checkValueB4);
            Assert.Equal(checkValueA5, checkValueB5);
            Assert.Equal(checkValueA6, checkValueB6);
            Assert.Equal(checkValueA7, checkValueB7);
            Assert.Equal(checkValueA8, checkValueB8);
            #endregion
        }
        #endregion

        #region private
        private string AppointmentsFindAll()
        {
            string returnValue = "";

            if (sqlController.AppointmentsFindOne(0) != null) returnValue += "0";
            if (sqlController.AppointmentsFindOne(1) != null) returnValue += "1";
            if (sqlController.AppointmentsFindOne(2) != null) returnValue += "2";
            if (sqlController.AppointmentsFindOne(3) != null) returnValue += "3";
            if (sqlController.AppointmentsFindOne(4) != null) returnValue += "4";

            if (sqlController.AppointmentsFindOne(WorkflowState.Canceled) != null) returnValue += "Canceled";
            if (sqlController.AppointmentsFindOne(WorkflowState.Completed) != null) returnValue += "Completed";
            if (sqlController.AppointmentsFindOne(WorkflowState.Created) != null) returnValue += "Created";
            if (sqlController.AppointmentsFindOne(WorkflowState.Failed_to_expection) != null) returnValue += "Failed_to_expection";
            if (sqlController.AppointmentsFindOne(WorkflowState.Failed_to_intrepid) != null) returnValue += "Failed_to_intrepid";
            if (sqlController.AppointmentsFindOne(WorkflowState.Planned) != null) returnValue += "Planned";
            if (sqlController.AppointmentsFindOne(WorkflowState.Processed) != null) returnValue += "Processed";
            if (sqlController.AppointmentsFindOne(WorkflowState.Retrived) != null) returnValue += "Retrived";
            if (sqlController.AppointmentsFindOne(WorkflowState.Revoked) != null) returnValue += "Revoked";
            if (sqlController.AppointmentsFindOne(WorkflowState.Sent) != null) returnValue += "Sent";

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
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<StartDate>", "</StartDate>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<EndDate>", "</EndDate>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<Language>", "</Language>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<Id>", "</Id>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<DefaultValue>", "</DefaultValue>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<MaxValue>", "</MaxValue>", "xxx", true);
            inputXmlString = t.ReplaceAtLocationAll(inputXmlString, "<MinValue>", "</MinValue>", "xxx", true);

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

            coreSdk_UT.CaseComplet(microtingUId, checkUId, workerMUId, unitMUId);
        }

        private void InteractionCaseComplet(int interactionCaseId)
        {
            var lst = sqlConSdk.UnitTest_FindAllActiveInteractionCaseLists(interactionCaseId);

            foreach (var item in lst)
            {
                CaseComplet(item.microting_uid, null);
            }
        }

        private string WaitForRestart()
        {
            int count = 0;
            while (count < 600)
            {
                if (!PrintLogLine().Contains("0:"))
                    break;

                Thread.Sleep(100);
                count++;
            }
            if (count == 600)
                throw new Exception("if (PrintLogLine().Contains(\"1:\")) failed 600 times");

            Thread.Sleep(1000);

            count = 0;
            while (count < 600)
            {
                if (coreOut.Running())
                    break;

                if (coreOut_UT.CoreDead())
                    break;

                Thread.Sleep(100);
                count++;
            }
            if (count == 600)
                throw new Exception("if (coreOut.Running()) failed 600 times");

            string rtrn = PrintLogLine();

            sqlController.UnitTest_TruncateTable(nameof(logs));
            sqlController.UnitTest_TruncateTable(nameof(log_exceptions));

            return rtrn;
        }

        private string PrintLogLine()
        {
            string str = "";
            str += sqlController.UnitTest_FindLog(1000, "Exception as per request");
            str += ":";
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountTried / Content:1");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountTried / Content:2");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountTried / Content:3");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountTried / Content:4");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountTried / Content:5");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountTried / Content:6");
            str += "/";
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountMax / Content:1");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountMax / Content:2");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountMax / Content:3");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountMax / Content:4");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountMax / Content:5");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:sameExceptionCountMax / Content:6");
            str += "/";
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:secondsDelay / Content:1");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:secondsDelay / Content:8");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:secondsDelay / Content:64");
            str += sqlController.UnitTest_FindLog(1000, "Variable Name:secondsDelay / Content:512");
            str += sqlController.UnitTest_FindLog(1000, "FatalExpection called for reason:'Restartfailed. Core failed to restart'");
            str += "/";
            str += sqlController.UnitTest_FindLog(1000, "Core triggered Exception event");
            str += Environment.NewLine;
            return str;
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
        public void EventException(object sender, EventArgs args)
        {
            lock (_lockFil)
            {
                sqlController.WriteLogEntry(new LogEntry(-4, "FATAL ERROR", "Core triggered Exception event"));
            }

            throw (Exception)sender; //Core need to be able that the external code crashed
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