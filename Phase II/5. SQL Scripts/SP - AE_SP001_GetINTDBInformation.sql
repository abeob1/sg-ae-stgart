ALTER Procedure [dbo].[AE_SP001_GetINTDBInformation]
@Entity as varchar(30)
as
begin
Declare @SQL varchar(max)

create table #FINAL 
(
    HTransID [int] NOT NULL,
	HOutlet [nvarchar](50) NOT NULL,
	HPOSTxNo [nvarchar](100) NOT NULL,
	HPOSTillId [nvarchar](100) NULL,
	PHOSTxDate [datetime]   NULL,
	HPOSTxDatetime [datetime]   NULL,
	[HPOSTxType] [nvarchar](5)   NULL,
	[HDepositNo] [nvarchar](50) NULL,
	[HSalesOrderNo] [nvarchar](50) NULL,
	[HGrossAmount] [numeric](19, 6) NULL,
	[HServiceCharge] [numeric](19, 6) NULL,
	[GST] [numeric](19, 6) NULL,
	[HGSTAdjustment] [numeric](19, 6) NULL,
	[HRounding] [numeric](19, 6) NULL,
	[HExcessAmount] [numeric](19, 6) NULL,
	[HTipsAmount] [numeric](19, 6) NULL,
	[HCovers] [numeric](19, 6) NULL,
	[HRevenueCenterName] [nvarchar](50) NULL,
	

	DTransID [int]   NULL,
	DHeaderID [nvarchar](50)   NULL,
	DOutlet [nvarchar](50)   NULL,
	DItemCode [nvarchar](100)   NULL,
	VatGourpSa [nvarchar] (20) NULL,
	DPriceBefDi [numeric](19, 6) NULL,
	DDiscPrcnt [numeric](19, 6)   NULL,
	DPrice [numeric](19, 6)   NULL,
	DQuantity [numeric](19, 6)  NULL,
	DLineTotal [numeric](19, 6) NULL,
	
	PPaymentAmount [numeric](19, 6)   NULL,
	DNetAmount [numeric](19, 6) NULL,
	[Validation2 Msg] [nvarchar](300) NULL,
	[Validation3 Msg] [nvarchar](300) NULL,
	[Validation4 Msg] [nvarchar](300) NULL,
	[Validation5 Msg] [nvarchar](300) NULL
	)

set @SQL  = '
SELECT T0.[ID] [HTransID],T0.[Outlet] [HOutlet],T0.[POSTxNo] [HPOSTxNo],T0.[POSTillId] [HPOSTillId],T0.[POSTxDate] [PHOSTxDate],
T0.[POSTxDatetime] [HPOSTxDatetime] ,T0.[POSTxType] [HPOSTxType] ,T0.[DepositNo] [HDepositNo],T0.[SaleOrderNo] [HSalesOrderNo], T0.[GrossAmount] [HGrossAmount]
,T0.[ServiceCharge] [HServiceCharge],T0.[GST],T0.[GSTAdjustment] [HGSTAdjustment],T0.[Rounding] [HRounding],T0.[ExcessAmount] [HExcessAmount] 
,T0.[TipsAmount] [HTipsAmount],T0.[Covers] [HCovers],T0.[RevenueCenterName] [HRevenueCenterName], 
T2.[ID] [DTransID],T2.[HeaderID] [DHeaderID] ,T2.[Outlet] [DOutlet],T2.[ItemCode] [DItemCode], T4.[VatGourpSa], T2.[PriceBefDi] [DPriceBefDi],
T2.[DiscPrcnt] [DDiscPrcnt],T2.[Price] [DPrice],T2.[Quantity] [DQuantity], T2.[LineTotal] [DLineTotal], 
(select sum(TT1.PaymentAmount)  From [AB_Payment] TT1 where TT1.HeaderID = T0.ID) [PPaymentAmount],
(select round(sum(TT.LineTotal + (TT.LineTotal * case when TT.Itemcode=''TIPS'' then 0 else 0.07 end)),2) + T0.Rounding  
from [AB_SalesTransDetail] TT 
where TT.HeaderID  = T0.ID 
 ) [DNetAmount],
case 
   when 
      isnull(T1.[U_POS_RefNo],'''') <> '''' then ''Receipt # '' + T1.[U_POS_RefNo] + '' already has an AR Invoice. {''+ cast(T1.DocNum as varchar) +''}'' 
   else '''' end [Validation2 Msg] ,
case 
  when 
      T4.SellItem = ''Y'' AND T5.U_AB_Inventory = ''Y'' AND T4.InvntItem = ''N''  AND ISNULL(T6.Code , '''') = '''' THEN ''Item Code '' + T2.[ItemCode] + '' has no sales BOM.'' 
  ELSE '''' END [Validation3 Msg] ,
case 
  when 
     (select round(sum(ISNULL(TT.LineTotal,0) + (ISNULL(TT.LineTotal,0) * case when TT.Itemcode=''TIPS'' then 0 else 0.07 end)),2) + ISNULL(T0.Rounding,0)
     from [AB_SalesTransDetail] TT 
     where TT.HeaderID  = T0.ID 
     ) <> (select isnull(sum(TT1.PaymentAmount),0)  From [AB_Payment] TT1 where TT1.HeaderID = T0.ID) then ''AR Invoice Total not equal to Payment Total.'' 
 else '''' end [Validation4 Msg],

 CASE WHEN (SELECT ISNULL(C.CardName,'''') FROM AB_Payment A INNER JOIN AB_SalesTransHeader B ON B.ID = A.HeaderID
			LEFT JOIN [Stuttgart LIVE].dbo.OCRC C ON C.CardName = A.PaymentCode
			WHERE A.HEADERID = T0.ID) = '''' THEN ''Payment Code does not exist in Credit Cards Setup.'' else '''' end [Validation5 Msg]

  FROM [AB_SalesTransHeader] T0 
left outer join ' + @Entity + '.. OINV T1 ON T1.[U_POS_RefNo] = T0.[POSTxNo] 
JOIN [AB_SalesTransDetail] T2 ON T2.HeaderID = T0.ID 
LEFT OUTER JOIN ' + @Entity + '.. OITM T4 ON T4.ITEMCODE = T2.ITEMCODE
JOIN ' + @Entity + '.. OITB T5 ON T5.ItmsGrpCod = T4.ItmsGrpCod 
LEFT OUTER JOIN ' + @Entity + '.. OITT T6 ON T6.Code = T2.ItemCode 
WHERE (isnull([Status], '''') = '''' OR [Status] =''FAIL'')
ORDER BY T0.ID , T2.ID '

insert into #FINAL 
 execute(@SQL)

SELECT #FINAL.HTransID , COUNT(#FINAL.[Validation2 Msg] )[Validation2]  into #Validation2 FROM #FINAL WHERE ISNULL(#FINAL.[Validation2 Msg],'') <> ''
GROUP BY #FINAL.HTransID

SELECT #FINAL.HTransID , COUNT(#FINAL.[Validation3 Msg] )[Validation3]  into #Validation3 FROM #FINAL WHERE ISNULL(#FINAL.[Validation3 Msg],'') <> ''
GROUP BY #FINAL.HTransID

SELECT #FINAL.HTransID , COUNT(#FINAL.[Validation4 Msg] )[Validation4]  into #Validation4 FROM #FINAL WHERE ISNULL(#FINAL.[Validation4 Msg],'') <> ''
GROUP BY #FINAL.HTransID

SELECT #FINAL.HTransID , COUNT(#FINAL.[Validation5 Msg] )[Validation5]  into #Validation5 FROM #FINAL WHERE ISNULL(#FINAL.[Validation5 Msg],'') <> ''
GROUP BY #FINAL.HTransID

SELECT #Final.* ,
CASE WHEN ISNULL(V2.Validation2 ,'') = '' THEN 0 ELSE V2.Validation2 END [Validation2Count],
CASE WHEN ISNULL(V3.Validation3 ,'') = '' THEN 0 ELSE V3.Validation3 END [Validation3Count],
CASE WHEN ISNULL(V4.Validation4 ,'') = '' THEN 0 ELSE V4.Validation4 END [Validation4Count],
CASE WHEN ISNULL(V5.Validation5 ,'') = '' THEN 0 ELSE V5.Validation5 END [Validation5Count],
ltrim(#final.[Validation2 Msg] + ' ' + #final.[Validation3 Msg] + ' ' + #final.[Validation4 Msg] + ' ' + #final.[Validation5 Msg]) [DetailsErrMsg]
 FROM #FINAL 
LEFT OUTER JOIN #Validation2 V2 ON V2.HTransID = #FINAL.HTransID
LEFT OUTER JOIN #Validation3 V3 ON V3.HTransID = #FINAL.HTransID
LEFT OUTER JOIN #Validation4 V4 ON V4.HTransID = #FINAL.HTransID
LEFT OUTER JOIN #Validation5 V5 ON V5.HTransID = #FINAL.HTransID
order by cast(#Final.DHeaderID as integer) , cast(#Final.DTransID as integer)

drop table #FINAL
drop table #Validation2
drop table #Validation3
drop table #Validation4
drop table #Validation5
End