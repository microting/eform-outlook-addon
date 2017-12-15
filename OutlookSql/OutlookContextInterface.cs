using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookSql
{
    interface OutlookContextInterface : IDisposable
    {
        DbSet<appointments> appointments { get; set; }
        DbSet<appointment_versions> appointment_versions { get; set; }
        DbSet<appointment_sites> appointment_sites { get; set; }
        DbSet<appointment_site_versions> appointment_site_versions { get; set; }
        DbSet<log_exceptions> log_exceptions { get; set; }
        DbSet<logs> logs { get; set; }
        DbSet<settings> settings { get; set; }

        int SaveChanges();

        Database Database { get; }
    }
}