﻿using Castle.MicroKernel.Registration;
using Castle.MicroKernel.SubSystems.Configuration;
using Castle.Windsor;
using Rebus.Config;
using System;


namespace Microting.OutlookAddon.Installers
{
    public class RebusInstaller : IWindsorInstaller
    {
        private readonly string connectionString;
        private readonly int maxParallelism;
        private readonly int numberOfWorkers;

        public RebusInstaller(string connectionString, int maxParallelism, int numberOfWorkers)
        {
            if (string.IsNullOrEmpty(connectionString)) throw new ArgumentNullException(nameof(connectionString));
            this.connectionString = connectionString;
            this.maxParallelism = maxParallelism;
            this.numberOfWorkers = numberOfWorkers;
        }

        public void Install(IWindsorContainer container, IConfigurationStore store)
        {
            Configure.With(new CastleWindsorContainerAdapter(container))
                .Logging(l => l.ColoredConsole())
                .Transport(t => t.UseSqlServer(connectionStringOrConnectionStringName: connectionString, tableName: "Rebus", inputQueueName: "eformoutlook-input"))
                .Options(o =>
                {
                    o.SetMaxParallelism(maxParallelism);
                    o.SetNumberOfWorkers(numberOfWorkers);
                })
                .Start();
        }
    }
}
