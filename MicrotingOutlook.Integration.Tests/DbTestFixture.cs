using NUnit.Framework;
using OutlookSql;
using System;
using System.Collections.Generic;
using System.Data.Entity.Core.Metadata.Edm;
using System.Data.Entity.Infrastructure;

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
            DbContext.Database.CommandTimeout = 300;

            DbContext.Database.CreateIfNotExists();

            DbContext.Database.Initialize(true);

            DoSetup();
        }

        [TearDown]
        public void TearDown()
        {
            var metadata = ((IObjectContextAdapter)DbContext).ObjectContext.MetadataWorkspace.GetItems(DataSpace.SSpace);

            List<string> tables = new List<string>();
            foreach (var item in metadata)
            {
                if (item.ToString().Contains("CodeFirstDatabaseSchema"))
                {
                    tables.Add(item.ToString().Replace("CodeFirstDatabaseSchema.", ""));
                }
            }

            foreach (string tableName in tables)
            {
                try
                {
                    DbContext.Database.ExecuteSqlCommand("DELETE FROM [" + tableName + "]");
                }
                catch
                {

                }

            }
            DbContext.Dispose();
        }

        public virtual void DoSetup() { }
    }
}
