namespace OutlookSql
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class OutlookDb : DbContext
    {
        public OutlookDb() { }

        public OutlookDb(string connectionString)
            : base(connectionString)
        {
        }

        public virtual DbSet<appointment_versions> appointment_versions { get; set; }
        public virtual DbSet<appointments> appointments { get; set; }
        public virtual DbSet<lookups> lookups { get; set; }
        public virtual DbSet<settings> settings { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<appointment_versions>()
                .Property(e => e.expectionString)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.workflow_state)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.created_at)
                .HasPrecision(0);

            modelBuilder.Entity<appointments>()
                .Property(e => e.updated_at)
                .HasPrecision(0);

            modelBuilder.Entity<appointments>()
                .Property(e => e.global_id)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.start_at)
                .HasPrecision(0);

            modelBuilder.Entity<appointments>()
                .Property(e => e.expire_at)
                .HasPrecision(0);

            modelBuilder.Entity<appointments>()
                .Property(e => e.subject)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.location)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.body)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.expectionString)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.site_ids)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.title)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.info)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.microting_uid)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.replacements)
                .IsUnicode(false);

            modelBuilder.Entity<settings>()
                .Property(e => e.name)
                .IsUnicode(false);

            modelBuilder.Entity<settings>()
                .Property(e => e.value)
                .IsUnicode(false);
        }
    }
}