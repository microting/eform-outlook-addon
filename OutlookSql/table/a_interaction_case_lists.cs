namespace OutlookSql
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class a_interaction_case_lists
    {
        public int id { get; set; }

        [StringLength(255)]
        public string workflow_state { get; set; }

        public int? version { get; set; }

        [Column(TypeName = "datetime2")]
        public DateTime? created_at { get; set; }

        [Column(TypeName = "datetime2")]
        public DateTime? updated_at { get; set; }

        public int? a_interaction_case_id { get; set; }

        public int? siteId { get; set; }

        public string stat { get; set; }

        public string microting_uid { get; set; }

        public string check_uid { get; set; }

        public int? case_id { get; set; }

        public virtual a_interaction_cases a_interaction_cases { get; set; }
    }
}