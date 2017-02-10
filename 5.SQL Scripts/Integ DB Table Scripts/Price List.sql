CREATE TABLE [dbo].[AB_PriceList](
	[ItemCode] [nvarchar](20) NULL,
	[PriceList] [int] NULL,
	[PriceListName] [nvarchar](50) NULL,
	[Currency] [nvarchar](10) NULL,
	[Price] [numeric](19, 6) NULL,
	[POSSyncDate] [datetime] NULL,
	[POSSyncDateTime] [datetime] NULL,
	[SAPSyncDate] [datetime] NULL,
	[SAPSyncDateTime] [datetime] NULL
)