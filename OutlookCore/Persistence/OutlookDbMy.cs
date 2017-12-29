namespace OutlookSql
{
    using System.Data.Entity;
    using MySql.Data.Entity;

    // Code-Based Configuration and Dependency resolution
    [DbConfigurationType(typeof(MySqlEFConfiguration))]
    public partial class OutlookDbMy : DbContext, OutlookContextInterface
    {
        public OutlookDbMy() { }

        public OutlookDbMy(string connectionString)
            : base(connectionString)
        {
        }

        public virtual DbSet<appointments> appointments { get; set; }
        public virtual DbSet<appointment_versions> appointment_versions { get; set; }
        public virtual DbSet<appointment_sites> appointment_sites { get; set; }
        public virtual DbSet<appointment_site_versions> appointment_site_versions { get; set; }
        public virtual DbSet<log_exceptions> log_exceptions { get; set; }
        public virtual DbSet<logs> logs { get; set; }
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
                .Property(e => e.processing_state)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.body)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.exceptionString)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.title)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.info)
                .IsUnicode(false);

            modelBuilder.Entity<appointments>()
                .Property(e => e.microting_uuid)
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