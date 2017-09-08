namespace OutlookSql
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class appointments
    {
        public int id { get; set; }

        [StringLength(255)]
        public string workflow_state { get; set; }

        public int? version { get; set; }

        [Column(TypeName = "datetime2")]
        public DateTime? created_at { get; set; }

        [Column(TypeName = "datetime2")]
        public DateTime? updated_at { get; set; }

        public string global_id { get; set; }

        [Column(TypeName = "datetime2")]
        public DateTime? start_at { get; set; }

        [Column(TypeName = "datetime2")]
        public DateTime? expire_at { get; set; }

        public int? duration { get; set; }

        [StringLength(255)]
        public string subject { get; set; }

        [StringLength(255)]
        public string location { get; set; }

        public string body { get; set; }

        public string expectionString { get; set; }

        public string site_ids { get; set; }

        [StringLength(255)]
        public string title { get; set; }

        [StringLength(255)]
        public string description { get; set; }

        public string info { get; set; }

        [StringLength(255)]
        public string microting_uid { get; set; }

        public short? connected { get; set; }

        public short? completed { get; set; }

        public string replacements { get; set; }

        public int? template_id { get; set; }

        public string response { get; set; }

        public short? color_rule { get; set; }
    }
}
