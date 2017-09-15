using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookSql
{
    interface MicrotingContextInterface : IDisposable
    {
        DbSet<a_interaction_case_lists> a_interaction_case_lists { get; set; }
        DbSet<a_interaction_cases> a_interaction_cases { get; set; }

        int SaveChanges();

        Database Database { get; }
    }
}