using OutlookSql;

using System;
using System.Collections.Generic;
using System.Linq;

using Xunit;

namespace UnitTest
{
    public class TestContext : IDisposable
    {
        bool useLiveData = false;

        //string connectionStringLocal_SDK_UnitTest = "Data Source=DESKTOP-7V1APE5\\SQLEXPRESS;Initial Catalog=MicrotingOutlookTest_SDK_UnitTest;Integrated Security=True"; //Uses unit test data
        //string connectionStringLocal_OUT_UnitTest = "Data Source=DESKTOP-7V1APE5\\SQLEXPRESS;Initial Catalog=MicrotingOutlookTest_OUT_UnitTest;Integrated Security=True"; //Uses unit test data

        string connectionStringLocal_SDK_UnitTest = "Persist Security Info=True;server=localhost;database=microtingMySQL_SDK;uid=root;password=1234";
        string connectionStringLocal_OUT_UnitTest = "Persist Security Info=True;server=localhost;database=microtingMySQL_OUT;uid=root;password=1234";

        string connectionStringLocal_SDK_LiveData = "Data Source=DESKTOP-7V1APE5\\SQLEXPRESS;Initial Catalog=MicrotingOutlookTest_SDK_LiveData;Integrated Security=True"; //Uses LIVE data
        string connectionStringLocal_OUT_LiveData = "Data Source=DESKTOP-7V1APE5\\SQLEXPRESS;Initial Catalog=MicrotingOutlookTest_OUT_LiveData;Integrated Security=True"; //Uses LIVE data

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
                    serverConnectionString_SDK = "Persist Security Info=True;server=localhost;database=microtingMySQL_SDK;uid=root;password="; //Uses travis database
                    serverConnectionString_OUT = "Persist Security Info=True;server=localhost;database=microtingMySQL_OUT;uid=root;password="; //Uses travis database
                    useLiveData = false;
                }
                else
                {
                    if (useLiveData)
                    {
                        serverConnectionString_SDK = connectionStringLocal_SDK_LiveData;
                        serverConnectionString_OUT = connectionStringLocal_OUT_LiveData;
                    }
                    else
                    {
                        serverConnectionString_SDK = connectionStringLocal_SDK_UnitTest;
                        serverConnectionString_OUT = connectionStringLocal_OUT_UnitTest;
                    }
                }
            }
            catch { }

                sqlCon = new SqlController(serverConnectionString_OUT);
            var sqlSdk = new eFormSqlController.SqlController(serverConnectionString_SDK);
            var adminT = new eFormCore.AdminTools(serverConnectionString_SDK);

            if (sqlSdk.SettingRead(eFormSqlController.Settings.token) == "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX")
                adminT.DbSetup("unittest");

            sqlCon.SettingUpdate(Settings.microtingDb, serverConnectionString_SDK);
        }
        #endregion

        #region once for all tests - teardown
        public void Dispose()
        {
            //sqlController.UnitTest_DeleteDb();
        }
        #endregion

        public string GetConnectionString()
        {
            return serverConnectionString_SDK;
        }

        public bool GetUseLiveData()
        {
            return useLiveData;
        }
        #endregion
    }

    [Collection("Database collection")]
    public class Outlook
    {
        [Fact]
        public void TestMethod1()
        {
            bool value = true;
            Assert.Equal(true, value);
        }
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