namespace OutlookSql
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class lookups
    {
        public int id { get; set; }
        
        public string title { get; set; }

        public string value { get; set; }
    }
}