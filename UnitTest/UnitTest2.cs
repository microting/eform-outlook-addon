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
    [Collection("Database collection")]
    public class UnitTest2
    {
        [Fact]
        public void Test000_Basics_0a_EnvironmentMachineName()
        {
            //Arrange
            string checkValueA = "";
            string checkValueB = "Not correct";

            //Act
            checkValueB = Environment.MachineName;

            //Assert
            Assert.Equal(checkValueA, checkValueB);
        }

        [Fact]
        public void Test000_Basics_1a_MustAlwaysPass()
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
}