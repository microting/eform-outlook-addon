namespace OutlookSql.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class ChangingMicrotingUidToSdkCaseIdOnAppointmentSites : DbMigration
    {
        public override void Up()
        {
            RenameColumn("dbo.appointment_sites", "microting_uuid", "sdk_case_id");
            RenameColumn("dbo.appointment_site_versions", "microting_uuid", "sdk_case_id");
        }
        
        public override void Down()
        {
            RenameColumn("dbo.appointment_sites", "sdk_case_id", "microting_uuid");
            RenameColumn("dbo.appointment_site_versions", "sdk_case_id", "microting_uuid");
        }
    }
}
