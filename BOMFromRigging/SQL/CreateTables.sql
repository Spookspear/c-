USE [Rigging7]
GO

/****** Object:  Table [dbo].[RiggingHeader]    Script Date: 20/09/2018 11:04:28 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[RiggingHeader](
	[WorksheetName]		[nvarchar](max) NULL,
	[WorkbookDate]		[datetime] NULL,
	[ContactPerson]		[nvarchar](50) NULL,
	[BudgetHolder]		[nvarchar](50) NULL,
	[VesselLocation]	[nvarchar](50) NULL,
	[ProjectDepartment] [nvarchar](50) NULL,
	[DateRequested]		[datetime] NULL,
	[ProjectDuration]	[nvarchar](50) NULL,
	[SAPCostCode]		[nvarchar](20) NULL,
	[DeliveryDetails]	[nvarchar](50) NULL,
	[Remarks]			[nvarchar](max) NULL,
	[ATRWONO]			[nvarchar](20) NULL,
	[Vendor]			[nvarchar](50) NULL,
	[PONumber]			[nvarchar](20) NULL,
	[LinkToLines]		[uniqueidentifier] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


CREATE TABLE [dbo].[RiggingLines](
	[WorksheetName]		[nvarchar](max) NULL,
	[HighLevelDesc]		[nvarchar](max) NULL,
	[LowLevelDesc]		[nvarchar](max) NULL,
	[Quantity]			[nvarchar](20) NULL,
	[QuantityDecimal]	[decimal](5, 0) NULL,
	[ItemValue]			[nvarchar](20) NULL,
	[ItemValueDecimal]	[decimal](10, 0) NULL,
	[TotalValue]		[nvarchar](20) NULL,
	[TotalValueDecimal]	[decimal](18, 0) NULL,
	[TestProcedure]		[nvarchar](50) NULL,
	[LineOrAdditional]	[nvarchar](1) NULL,
	[LinkToHeader]		[uniqueidentifier] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO


