namespace RiggingConsoleApp.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class additionfieldsandlines : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.RiggingLineDS",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        HighLevelDesc = c.String(),
                        LowLevelDesc = c.String(),
                        Quantity = c.String(),
                        ItemValue = c.String(),
                        TotalValue = c.String(),
                        TestProcedure = c.String(),
                        LineOrAdditional = c.String(),
                        RiggingHeaderId = c.Int(nullable: false),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.RiggingHeaderDS", t => t.RiggingHeaderId, cascadeDelete: true)
                .Index(t => t.RiggingHeaderId);
            
            AddColumn("dbo.RiggingHeaderDS", "FileDate", c => c.DateTime(nullable: false));
            AddColumn("dbo.RiggingHeaderDS", "ContactPerson", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "BudgetHolder", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "VesselLocation", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "ProjectDepartment", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "DateRequested", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "DateRequired", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "ProjectDuration", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "SAPCostCode", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "DeliveryDetails", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "Remarks", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "ATRWONO", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "Vendor", c => c.String());
            AddColumn("dbo.RiggingHeaderDS", "PONumber", c => c.String());
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.RiggingLineDS", "RiggingHeaderId", "dbo.RiggingHeaderDS");
            DropIndex("dbo.RiggingLineDS", new[] { "RiggingHeaderId" });
            DropColumn("dbo.RiggingHeaderDS", "PONumber");
            DropColumn("dbo.RiggingHeaderDS", "Vendor");
            DropColumn("dbo.RiggingHeaderDS", "ATRWONO");
            DropColumn("dbo.RiggingHeaderDS", "Remarks");
            DropColumn("dbo.RiggingHeaderDS", "DeliveryDetails");
            DropColumn("dbo.RiggingHeaderDS", "SAPCostCode");
            DropColumn("dbo.RiggingHeaderDS", "ProjectDuration");
            DropColumn("dbo.RiggingHeaderDS", "DateRequired");
            DropColumn("dbo.RiggingHeaderDS", "DateRequested");
            DropColumn("dbo.RiggingHeaderDS", "ProjectDepartment");
            DropColumn("dbo.RiggingHeaderDS", "VesselLocation");
            DropColumn("dbo.RiggingHeaderDS", "BudgetHolder");
            DropColumn("dbo.RiggingHeaderDS", "ContactPerson");
            DropColumn("dbo.RiggingHeaderDS", "FileDate");
            DropTable("dbo.RiggingLineDS");
        }
    }
}
