namespace OutlookSql
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class appointment_site_versions
    {
        [Key]
        public int id { get; set; }

        public int? appointment_site_id { get; set; }

        [StringLength(255)]
        public string workflow_state { get; set; }

        public int? version { get; set; }

        public DateTime? created_at { get; set; }

        public DateTime? updated_at { get; set; }

        public string exceptionString { get; set; }

        [StringLength(255)]
        public string microting_uuid { get; set; }

        [StringLength(255)]
        public string processing_state { get; set; }

        public short? completed { get; set; }

        public virtual appointments appointment { get; set; }
    }
}
