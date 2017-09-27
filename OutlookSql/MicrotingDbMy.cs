namespace OutlookSql
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class MicrotingDbMy : DbContext, MicrotingContextInterface
    {
        public MicrotingDbMy() { }

        public MicrotingDbMy(string connectionString)
             : base(connectionString)
        {
        }

        public virtual DbSet<a_interaction_case_lists>  a_interaction_case_lists { get; set; }
        public virtual DbSet<a_interaction_cases>       a_interaction_cases { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            //modelBuilder.Entity<a_interaction_cases>()
            //    .HasMany(e => e.a_interaction_case_lists)
            //    .WithOptional(e => e.a_interaction_cases)
            //    .HasForeignKey(e => e.a_interaction_case_id);
        }
    }
}
