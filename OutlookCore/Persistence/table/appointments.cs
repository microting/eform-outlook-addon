namespace OutlookSql
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class appointments
    {

        public appointments()
        {
            this.appointment_sites = new HashSet<appointment_sites>();
        }
        [Key]
        public int id { get; set; }

        [StringLength(255)]
        public string workflow_state { get; set; }

        public int? version { get; set; }

        public DateTime? created_at { get; set; }

        public DateTime? updated_at { get; set; }

        public string global_id { get; set; }

        public DateTime? start_at { get; set; }

        public DateTime? expire_at { get; set; }

        public int? duration { get; set; }

        [StringLength(255)]
        public string subject { get; set; }

        [StringLength(255)]
        public string processing_state { get; set; }

        public string body { get; set; }

        public string exceptionString { get; set; }

        [StringLength(255)]
        public string title { get; set; }

        [StringLength(255)]
        public string description { get; set; }

        public string info { get; set; }

        [StringLength(255)]
        public string microting_uuid { get; set; }

        public short? completed { get; set; }

        public string replacements { get; set; }

        public int? template_id { get; set; }

        public string response { get; set; }

        public short? color_rule { get; set; }

        public virtual ICollection<appointment_sites> appointment_sites { get; set; }

        public override string ToString()
        {
            string globalId = "";
            string start = "";
            string _title = "";
            string _processing_state = "";

            if (global_id != null)
                globalId = global_id;

            if (start_at != null)
                start = start_at.ToString();

            if (title != null)
                _title = title;

            if (processing_state != null)
                _processing_state = processing_state;

            return "GlobalId:" + globalId + " / Start:" + start + " / Title:" + _title + " / Processing state:" + _processing_state;
        }
    }
}