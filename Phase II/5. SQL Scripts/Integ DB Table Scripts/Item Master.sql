CREATE TABLE [dbo].[AB_ItemMaster](
	[ItemCode] [nvarchar](20) NULL,
	[ItemName] [nvarchar](100) NULL,
	[FrgnName] [nvarchar](30) NULL,
	[EASIGroup] [nvarchar](15) NULL,
	[EASIDept] [nvarchar](15) NULL,
	[ProductType] [nvarchar](2) NULL,
	[SalesUnitMsr] [nvarchar](15) NULL,
	[Barcode] [nvarchar](16) NULL,
	[Active] [nvarchar](1) NULL,
	[ServiceCharge] [nvarchar](1) NULL,
	[GST] [nvarchar](1) NULL,
	[AllowDiscount] [nvarchar](1) NULL,
	[AllowZero] [nvarchar](1) NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
)