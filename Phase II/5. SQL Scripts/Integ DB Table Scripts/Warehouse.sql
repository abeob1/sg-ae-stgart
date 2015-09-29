CREATE TABLE [dbo].[AB_Warehouses](
	[WhsCode] [nvarchar](8) NULL,
	[WhsName] [nvarchar](50) NULL,
	[Active] [nvarchar](1) NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
)