namespace OutlookSql.Migrations
{
    using System;
    using System.Data.Entity.Migrations;

    public partial class InitialDb : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.appointment_site_versions",
                c => new
                {
                    id = c.Int(nullable: false, identity: true),
                    appointment_id = c.Int(),
                    microting_site_uid = c.Int(nullable: false),
                    workflow_state = c.String(maxLength: 255),
                    version = c.Int(),
                    created_at = c.DateTime(),
                    updated_at = c.DateTime(),
                    exceptionString = c.String(),
                    microting_uuid = c.String(maxLength: 255),
                    processing_state = c.String(maxLength: 255),
                    completed = c.Short(),
                    appointment_site_id = c.Int(),
                })
                .PrimaryKey(t => t.id);

            CreateTable(
                "dbo.appointment_sites",
                c => new
                {
                    id = c.Int(nullable: false, identity: true),
                    appointment_id = c.Int(),
                    microting_site_uid = c.Int(nullable: false),
                    workflow_state = c.String(maxLength: 255),
                    version = c.Int(),
                    created_at = c.DateTime(),
                    updated_at = c.DateTime(),
                    exceptionString = c.String(),
                    microting_uuid = c.String(maxLength: 255),
                    processing_state = c.String(maxLength: 255),
                    completed = c.Short(),
                })
                .PrimaryKey(t => t.id)
                .ForeignKey("dbo.appointments", t => t.appointment_id)
                .Index(t => t.appointment_id);

            CreateTable(
                "dbo.appointments",
                c => new
                {
                    id = c.Int(nullable: false, identity: true),
                    workflow_state = c.String(maxLength: 255, unicode: false),
                    version = c.Int(),
                    created_at = c.DateTime(),
                    updated_at = c.DateTime(),
                    global_id = c.String(unicode: false),
                    start_at = c.DateTime(),
                    expire_at = c.DateTime(),
                    duration = c.Int(),
                    subject = c.String(maxLength: 255, unicode: false),
                    processing_state = c.String(maxLength: 255, unicode: false),
                    body = c.String(unicode: false),
                    exceptionString = c.String(unicode: false),
                    title = c.String(maxLength: 255, unicode: false),
                    description = c.String(maxLength: 255),
                    info = c.String(unicode: false),
                    microting_uuid = c.String(maxLength: 255, unicode: false),
                    completed = c.Short(),
                    replacements = c.String(unicode: false),
                    template_id = c.Int(),
                    response = c.String(),
                    color_rule = c.Short(),
                })
                .PrimaryKey(t => t.id);

            CreateTable(
                "dbo.appointment_versions",
                c => new
                {
                    id = c.Int(nullable: false, identity: true),
                    appointment_id = c.Int(),
                    workflow_state = c.String(maxLength: 255),
                    version = c.Int(),
                    created_at = c.DateTime(),
                    updated_at = c.DateTime(),
                    global_id = c.String(),
                    start_at = c.DateTime(),
                    expire_at = c.DateTime(),
                    duration = c.Int(),
                    subject = c.String(maxLength: 255),
                    location = c.String(maxLength: 255),
                    body = c.String(),
                    expectionString = c.String(unicode: false),
                    site_ids = c.String(),
                    title = c.String(maxLength: 255),
                    description = c.String(maxLength: 255),
                    info = c.String(),
                    microting_uid = c.String(maxLength: 255),
                    connected = c.Short(),
                    completed = c.Short(),
                    replacements = c.String(),
                    template_id = c.Int(),
                    response_text = c.String(),
                    color_rule = c.Short(),
                })
                .PrimaryKey(t => t.id);

            CreateTable(
                "dbo.log_exceptions",
                c => new
                {
                    id = c.Int(nullable: false, identity: true),
                    created_at = c.DateTime(nullable: false),
                    level = c.Int(nullable: false),
                    type = c.String(),
                    message = c.String(),
                })
                .PrimaryKey(t => t.id);

            CreateTable(
                "dbo.logs",
                c => new
                {
                    id = c.Int(nullable: false, identity: true),
                    created_at = c.DateTime(nullable: false),
                    level = c.Int(nullable: false),
                    type = c.String(),
                    message = c.String(),
                })
                .PrimaryKey(t => t.id);

            CreateTable(
                "dbo.settings",
                c => new
                {
                    id = c.Int(nullable: false),
                    name = c.String(nullable: false, maxLength: 50, unicode: false),
                    value = c.String(unicode: false),
                })
                .PrimaryKey(t => t.id);

        }

        public override void Down()
        {
            DropForeignKey("dbo.appointment_sites", "appointment_id", "dbo.appointments");
            DropIndex("dbo.appointment_sites", new[] { "appointment_id" });
            DropTable("dbo.settings");
            DropTable("dbo.logs");
            DropTable("dbo.log_exceptions");
            DropTable("dbo.appointment_versions");
            DropTable("dbo.appointments");
            DropTable("dbo.appointment_sites");
            DropTable("dbo.appointment_site_versions");
        }
    }
}
