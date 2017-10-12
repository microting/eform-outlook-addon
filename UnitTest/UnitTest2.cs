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
    public class UnitTest2
    {
        [Fact]
        public void Test001()
        {
            //Arrange
            bool checkValueA = true;
            bool checkValueB = false;

            //Act
            checkValueB = true;

            //Assert
            Assert.Equal(checkValueA, checkValueB);
        }

        //[Fact]
        //public void Test002()
        //{
        //    //Arrange
        //    string checkValueA = "";
        //    string checkValueB = "Panic";

        //    //Act
        //    checkValueB = Environment.MachineName;

        //    //Assert
        //    Assert.Equal(checkValueA, checkValueB);
        //}

        [Fact]
        public void Test003()
        {
            //Arrange
            string checkValueA = "";
            string checkValueB = Environment.MachineName;
            string connectionString = "Data Source=(localdb)\\v11.0;Initial Catalog=" + "UnitTest_Outlook_" + "Microting" + ";Integrated Security=SSPI"; //vsts database

            //Act
            SqlController sql = new SqlController(connectionString);
            checkValueA = "yes";
            checkValueB = "yes";

            //Assert
            Assert.Equal(checkValueA, checkValueB);
        }
    }
}