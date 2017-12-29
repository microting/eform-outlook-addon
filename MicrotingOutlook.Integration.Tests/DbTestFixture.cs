using NUnit.Framework;
using OutlookSql;

namespace MicrotingOutlook.Integration.Tests
{
    [TestFixture]
    public abstract class DbTestFixture
    {
        protected OutlookDbMs DbContext;
        protected string ConnectionString => @"data source=(LocalDb)\SharedInstance;Initial catalog=eformoutlook-tests";

        [SetUp]
        public void Setup()
        {
            DbContext = new OutlookDbMs(ConnectionString);
            DbContext.Database.CreateIfNotExists();

            DbContext.Database.Initialize(false);

            DoSetup();
        }

        [TearDown]
        public void TearDown()
        {
            DbContext.Database.Delete();
            DbContext.Dispose();
        }

        public virtual void DoSetup() { }
    }
}
