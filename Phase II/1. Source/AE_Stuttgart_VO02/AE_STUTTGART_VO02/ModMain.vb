Module ModMain

    Public oDT_BOM As DataTable = Nothing
    Public oDT_InvoiceData As DataTable = Nothing
    Public oDT_PaymentData As DataTable = Nothing

    Sub Main()

        Dim sFuncName As String = String.Empty
        Dim sErrDesc As String = String.Empty
        ' Dim oARInvoice As AE_STUTTGART_DLL.clsARInvoice = New AE_STUTTGART_DLL.clsARInvoice

        Dim sQuery As String = String.Empty

        Try
            p_iDebugMode = DEBUG_ON
            sFuncName = "Main()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            'Getting the Parameter Values from App Cofig File
            Console.WriteLine("Calling GetSystemIntializeInfo() ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling GetSystemIntializeInfo()", sFuncName)
            If GetSystemIntializeInfo(p_oCompDef, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If Not oDT_InvoiceData Is Nothing Then
                '' Function to connect the Company
                If p_oCompany Is Nothing Then
                    Console.WriteLine("Calling ConnectToTargetCompany() ", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
                    If ConnectToTargetCompany(p_oCompany, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                End If

                'Console.WriteLine("Calling IntegrityValidation() ", sFuncName)
                'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling IntegrityValidation()", sFuncName)
                'If IntegrityValidation(oDT_InvoiceData, oDT_PaymentData, p_oCompany, sErrDesc) <> RTN_SUCCESS Then
                '    Call WriteToLogFile(sErrDesc, sFuncName)
                '    Console.WriteLine("Completed with ERROR : " & sErrDesc, sFuncName)
                '    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                'End If
            Else

                Console.WriteLine("There is No Pending Records Found in Integration DB", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("There is No Pending Records Found in Integration DB", sFuncName)
            End If

            '************MASTER DATA SYNC CODE ON 16/09/2015 STARTS**************************
            '************ITEM MASTER SYNC CODE ON 16/09/2015 STARTS**************************
            Console.WriteLine("Item Master sync Starts", sFuncName)
            sQuery = "INSERT INTO [" & p_oCompDef.p_sIntDBName & "].dbo.AB_ItemMaster(ItemCode,ItemName,FrgnName,EASIGroup,EASIDept,ProductType,SalesUnitMsr,Barcode,Active,ServiceCharge,GST,AllowDiscount,AllowZero,SAPSyncDate,SAPSyncDateTime)"
            sQuery = sQuery & " SELECT T0.ItemCode,T0.ItemName,T0.U_POSDesc,T0.ItmsGrpCod,T0.U_AB_EASIDept, 'SS' [ProductType],T0.SalUnitMsr,T0.CodeBars,T0.validFor,"
            sQuery = sQuery & " CASE WHEN ISNULL(T0.U_POSNoService,'FALSE') = 'FALSE' THEN 1 ELSE 0 END [U_POSNoService],"
            sQuery = sQuery & " CASE WHEN ISNULL(T0.U_POSNoGst,'FALSE') = 'FALSE' THEN 1 ELSE 0 END [U_POSNoGst],"
            sQuery = sQuery & " CASE WHEN ISNULL(T0.U_POSNonDisc,'FALSE') = 'FALSE' THEN 1 ELSE 0 END [U_POSNonDisc],"
            sQuery = sQuery & " CASE WHEN ISNULL(T0.U_POSZeroPrice,'FALSE') = 'FALSE' THEN 0 ELSE 1 END  [U_POSZeroPrice],"
            sQuery = sQuery & " DATEADD(DAY, DATEDIFF(day, 0, GETDATE()), 0) [SAPSyncDate],GETDATE() [SAPSyncDateTime]"
            sQuery = sQuery & " FROM [" & p_oCompDef.p_sDataBaseName & "].dbo.OITM T0 "
            sQuery = sQuery & " LEFT JOIN [" & p_oCompDef.p_sDataBaseName & "].dbo.OITB T1 ON T1.ItmsGrpCod = T0.ItmsGrpCod "
            sQuery = sQuery & " WHERE T0.ItemCode NOT IN (SELECT ItemCode FROM [" & p_oCompDef.p_sIntDBName & "].dbo.AB_ItemMaster) "

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item Master Sync Query Exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while synchoronizing with item master", sFuncName)
            Else
                Console.WriteLine("Item master synchronization completed Successfully", sFuncName)
            End If

            '************ITEM MASTER UPDATE CODE**************************
            Console.WriteLine("Update Item Master Datas", sFuncName)
            sQuery = "UPDATE [" & p_oCompDef.p_sIntDBName & "].dbo.AB_ItemMaster"
            sQuery = sQuery & " SET ItemName = T1.ItemName,FrgnName = T1.U_POSDesc,EASIGroup = T1.ItmsGrpCod, "
            sQuery = sQuery & " EASIDept = T1.U_AB_EASIDept,ProductType = T1.ProductType,SalesUnitMsr = T1.SalUnitMsr, Barcode = T1.CodeBars, "
            sQuery = sQuery & " Active = T1.ValidFor,ServiceCharge = T1.U_POSNoService,GST = T1.U_POSNoGst, "
            sQuery = sQuery & " AllowDiscount = T1.U_POSNonDisc,AllowZero = T1.U_POSZeroPrice, "
            sQuery = sQuery & " SAPSyncDate = DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),0),SAPSyncDateTime = GETDATE() "
            sQuery = sQuery & " FROM [" & p_oCompDef.p_sIntDBName & "].dbo.AB_ItemMaster T0"
            sQuery = sQuery & " INNER JOIN (SELECT T0.ItemCode,T0.ItemName,T0.U_POSDesc,T0.ItmsGrpCod,T0.U_AB_EASIDept, 'SS' [ProductType],T0.SalUnitMsr,T0.CodeBars,T0.ValidFor,"
            sQuery = sQuery & " 		    CASE WHEN ISNULL(T0.U_POSNoService,'FALSE') = 'FALSE' THEN 1 ELSE 0 END [U_POSNoService],"
            sQuery = sQuery & " 			CASE WHEN ISNULL(T0.U_POSNoGst,'FALSE') = 'FALSE' THEN 1 ELSE 0 END [U_POSNoGst],"
            sQuery = sQuery & " 			CASE WHEN ISNULL(T0.U_POSNonDisc,'FALSE') = 'FALSE' THEN 1 ELSE 0 END [U_POSNonDisc],"
            sQuery = sQuery & " 			CASE WHEN ISNULL(T0.U_POSZeroPrice,'FALSE') = 'FALSE' THEN 0 ELSE 1 END  [U_POSZeroPrice] "
            sQuery = sQuery & " 			FROM [" & p_oCompDef.p_sDataBaseName & "].dbo.OITM T0"
            sQuery = sQuery & " 			INNER JOIN [" & p_oCompDef.p_sDataBaseName & "].dbo.OITB T1 ON T1.ItmsGrpCod = T0.ItmsGrpCod"
            sQuery = sQuery & " 			WHERE T0.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- " & p_oCompDef.p_iIntegDays & ")"
            sQuery = sQuery & " 		    ) T1 ON T1.ItemCode = T0.ItemCode"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item Master Sync update Query Exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while updating item master", sFuncName)
            Else
                Console.WriteLine("Item master update Successful", sFuncName)
            End If

            Console.WriteLine("Update Item Master Active Status", sFuncName)
            sQuery = "UPDATE [" & p_oCompDef.p_sIntDBName & "].dbo.AB_ItemMaster"
            sQuery = sQuery & " SET Active = 'N'"
            sQuery = sQuery & " WHERE ItemCode IN (SELECT SWW FROM [" & p_oCompDef.p_sDataBaseName & "].dbo.OITM)"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item Master Active Status update Query Exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while updating item master Active Status", sFuncName)
            Else
                Console.WriteLine("Item master Active Status update Successful", sFuncName)
            End If

            '************EASI DEPARTMENT SYNC CODE ON 16/09/2015 STARTS**************************
            Console.WriteLine("EASI Department Sync starts", sFuncName)
            sQuery = "DELETE FROM [" & p_oCompDef.p_sIntDBName & "].dbo.AB_EASIDepartment"
            sQuery = sQuery & " INSERT INTO [" & p_oCompDef.p_sIntDBName & "].dbo.AB_EASIDepartment(Code,Name,SAPSyncDate,SAPSyncDateTime)"
            sQuery = sQuery & " SELECT Code,Name,DATEADD(DAY, DATEDIFF(day, 0, GETDATE()), 0) [SAPSyncDate],GETDATE() [SAPSyncDateTime] "
            sQuery = sQuery & " FROM [" & p_oCompDef.p_sDataBaseName & "].dbo.[@AB_EASIDEPT] T0 "
            sQuery = sQuery & " WHERE T0.Code NOT IN (SELECT Code FROM  [" & p_oCompDef.p_sIntDBName & "].dbo.AB_EASIDepartment) "

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item Group Sync Query Exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while synchoronizing EASI Department", sFuncName)
            Else
                Console.WriteLine("EASI Department synchronization completed Successfully", sFuncName)
            End If

            '******************************EASI GROUP SYNC CODE ON 23/09/2015*****************
            Console.WriteLine("EASI Group Sync starts", sFuncName)
            sQuery = " INSERT INTO [" & p_oCompDef.p_sIntDBName & "].dbo.AB_EASIGroup(Code,Name,SAPSyncDate,SAPSyncDateTime)"
            sQuery = sQuery & " SELECT T0.ItmsGrpCod,T0.ItmsGrpNam,DATEADD(DAY, DATEDIFF(day, 0, GETDATE()), 0) [SAPSyncDate],GETDATE() [SAPSyncDateTime]"
            sQuery = sQuery & " FROM [" & p_oCompDef.p_sDataBaseName & "].dbo.OITB T0 "
            sQuery = sQuery & " WHERE T0.ItmsGrpCod NOT IN (SELECT Code FROM [" & p_oCompDef.p_sIntDBName & "].dbo.AB_EASIGROUP)"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EASI Group Sync Query Exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while synchoronizing EASI Group", sFuncName)
            Else
                Console.WriteLine("EASI Group synchronization completed Successfully", sFuncName)
            End If

            '**********UPDATE EASI Group CODE******************
            Console.WriteLine("Update EASI Group starts", sFuncName)
            sQuery = "UPDATE [" & p_oCompDef.p_sIntDBName & "].dbo.AB_EASIGroup "
            sQuery = sQuery & " SET Code = T1.ItmsGrpCod,Name = T1.ItmsGrpNam,SAPSyncDate = DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),0),SAPSyncDateTime = GETDATE()"
            sQuery = sQuery & " FROM [" & p_oCompDef.p_sIntDBName & "].dbo.AB_EASIGroup T0"
            sQuery = sQuery & " INNER JOIN (SELECT T0.ItmsGrpCod,T0.ItmsGrpNam"
            sQuery = sQuery & " 			FROM [" & p_oCompDef.p_sDataBaseName & "].dbo.OITB T0"
            sQuery = sQuery & " 			WHERE T0.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- " & p_oCompDef.p_iIntegDays & ")"
            sQuery = sQuery & " 		  ) T1 ON T1.ItmsGrpCod = T0.Code"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item Group Update Query Exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while updating EASI Group", sFuncName)
            Else
                Console.WriteLine("EASI Group update Successful", sFuncName)
            End If

            '************PRICE LIST SYNC CODE ON 16/09/2015 STARTS**************************
            Console.WriteLine("Price List sync starts", sFuncName)
            sQuery = "INSERT INTO [" & p_oCompDef.p_sIntDBName & "].dbo.AB_PriceList(ItemCode,PriceList,PriceListName,Currency,Price,SAPSyncDate,SAPSyncDateTime)"
            sQuery = sQuery & " SELECT T0.ItemCode,T2.ListNum,T2.ListName ,T1.Currency,  "
            sQuery = sQuery & " CASE WHEN T0.UgpEntry = -1 THEN T1.Price "
            sQuery = sQuery & "      WHEN (T0.UgpEntry <> -1 AND T0.SUoMEntry =  T0.IUoMEntry) THEN T1.Price "
            sQuery = sQuery & "      ELSE T3.Price END [Price],"
            sQuery = sQuery & " DATEADD(DAY, DATEDIFF(day, 0, GETDATE()), 0) [SAPSyncDate],GETDATE() [SAPSyncDateTime] "
            sQuery = sQuery & " FROM [" & p_oCompDef.p_sDataBaseName & "].dbo.OITM T0 "
            sQuery = sQuery & " LEFT JOIN [" & p_oCompDef.p_sDataBaseName & "].dbo.ITM1 T1 ON T0.ItemCode = T1.ItemCode AND T1.PriceList = 1 "
            sQuery = sQuery & " LEFT JOIN [" & p_oCompDef.p_sDataBaseName & "].dbo.ITM9 T3 ON T0.ItemCode = T3.ItemCode AND T3.PriceList = 1 AND T0.SUoMEntry = T3.UomEntry "
            sQuery = sQuery & " LEFT JOIN [" & p_oCompDef.p_sDataBaseName & "].dbo.OPLN T2 ON T2.ListNum = T1.PriceList "
            sQuery = sQuery & " WHERE T0.ItemCode NOT IN (SELECT ItemCode FROM [" & p_oCompDef.p_sIntDBName & "].dbo.AB_PriceList) "
            sQuery = sQuery & " AND (T1.PriceList = 1 OR T3.PriceList = 1) "

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Price list Sync Query Exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while synchoronizing Price list", sFuncName)
            Else
                Console.WriteLine("Price list synchronization completed Successfully", sFuncName)
            End If

            '***********PRICE LIST UPDATE*******************
            Console.WriteLine("Update Price list starts", sFuncName)
            sQuery = "UPDATE [" & p_oCompDef.p_sIntDBName & "].dbo.AB_PriceList"
            sQuery = sQuery & " SET PriceListName = T1.ListName,Currency = T1.Currency,Price = T1.Price, "
            sQuery = sQuery & " SAPSyncDate = DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),0),SAPSyncDateTime = GETDATE() "
            sQuery = sQuery & " FROM [" & p_oCompDef.p_sIntDBName & "].dbo.AB_PriceList T0"
            sQuery = sQuery & " INNER JOIN (SELECT T0.ItemCode,T2.ListNum,T2.ListName ,T1.Currency,"
            sQuery = sQuery & " 			CASE WHEN T0.UgpEntry = -1 THEN T1.Price "
            sQuery = sQuery & "                  WHEN (T0.UgpEntry <> -1 AND T0.SUoMEntry =  T0.IUoMEntry) THEN T1.Price"
            sQuery = sQuery & "                  ELSE T3.Price END [Price]"
            sQuery = sQuery & " 			FROM [" & p_oCompDef.p_sDataBaseName & "].dbo.OITM T0"
            sQuery = sQuery & " 			LEFT JOIN [" & p_oCompDef.p_sDataBaseName & "].dbo.ITM1 T1 ON T0.ItemCode = T1.ItemCode AND T1.PriceList = 1"
            sQuery = sQuery & " 			LEFT JOIN [" & p_oCompDef.p_sDataBaseName & "].dbo.OPLN T2 ON T2.ListNum = T1.PriceList"
            sQuery = sQuery & " 			LEFT JOIN [" & p_oCompDef.p_sDataBaseName & "].dbo.ITM9 T3 ON T3.ItemCode = T0.ItemCode AND T3.PriceList = 1 AND T0.SUoMEntry = T3.UomEntry"
            sQuery = sQuery & " 			WHERE T0.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- " & p_oCompDef.p_iIntegDays & ")"
            sQuery = sQuery & " 		   ) T1 ON T1.ItemCode = T0.ItemCode AND T1.ListNum = T0.PriceList "

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Price list update Query Exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while updating Price list", sFuncName)
            Else
                Console.WriteLine("Price list update Successful", sFuncName)
            End If

            '************WAREHOUSE SYNC CODE ON 16/09/2015 STARTS**************************
            Console.WriteLine("Warehouse sync starts", sFuncName)
            sQuery = "INSERT INTO [" & p_oCompDef.p_sIntDBName & "].dbo.AB_Warehouses(WhsCode,WhsName,Active,SAPSyncDate,SAPSyncDateTime)"
            sQuery = sQuery & " SELECT WhsCode,WhsName,CASE WHEN Inactive = 'N' THEN 'Y' ELSE 'N' END [Active], "
            sQuery = sQuery & " DATEADD(DAY, DATEDIFF(day, 0, GETDATE()), 0) [SAPSyncDate],GETDATE() [SAPSyncDateTime]"
            sQuery = sQuery & " FROM [" & p_oCompDef.p_sDataBaseName & "].dbo.OWHS T0 "
            sQuery = sQuery & " WHERE T0.WhsCode NOT IN (SELECT WhsCode FROM [" & p_oCompDef.p_sIntDBName & "].dbo.AB_Warehouses) "

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Warehouse Sync Query Exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while synchoronizing warehouses", sFuncName)
            Else
                Console.WriteLine("Warehouses synchronization completed Successfully", sFuncName)
            End If

            '***************UPDATE WAREHOUSE***************
            Console.WriteLine("Update Warehouse starts", sFuncName)
            sQuery = "UPDATE [" & p_oCompDef.p_sIntDBName & "].dbo.AB_Warehouses"
            sQuery = sQuery & " SET WhsName = T1.WhsName,SAPSyncDate = DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),0),SAPSyncDateTime = GETDATE()"
            sQuery = sQuery & " FROM [" & p_oCompDef.p_sIntDBName & "].dbo.AB_Warehouses T0"
            sQuery = sQuery & " INNER JOIN [" & p_oCompDef.p_sDataBaseName & "].dbo.OWHS T1 ON T1.WhsCode = T0.WhsCode"
            sQuery = sQuery & " WHERE T1.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- " & p_oCompDef.p_iIntegDays & ")"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Warehouse Updation Query Exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while updating warehouses", sFuncName)
            Else
                Console.WriteLine("Warehouses update Successful", sFuncName)
            End If

            '************TENDER MODE SYNC CODE ON 21/09/2015 STARTS**************************
            Console.WriteLine("Tender sync starts", sFuncName)
            sQuery = "INSERT INTO [" & p_oCompDef.p_sIntDBName & "].dbo.AB_Tender(CreditCard,CardName,SAPSyncDate,SAPSyncDateTime)"
            sQuery = sQuery & " SELECT CreditCard,CardName,DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),0) [SAPSyncDate],GETDATE() [SAPSyncDateTime] "
            sQuery = sQuery & " FROM [" & p_oCompDef.p_sDataBaseName & "].dbo.OCRC T0 "
            sQuery = sQuery & " WHERE T0.CreditCard NOT IN (SELECT CreditCard FROM [" & p_oCompDef.p_sIntDBName & "].dbo.AB_Tender) "

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Tender sync Query exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while synchronizing Tender", sFuncName)
            Else
                Console.WriteLine("Tender synchronization Successful", sFuncName)
            End If

            '***************UPDATE TENDER***************
            Console.WriteLine("Update Tender starts", sFuncName)
            sQuery = "UPDATE [" & p_oCompDef.p_sIntDBName & "].dbo.AB_Tender"
            sQuery = sQuery & " SET CreditCard = T1.CreditCard,SAPSyncDate = DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),0),SAPSyncDateTime = GETDATE() "
            sQuery = sQuery & " FROM AB_Tender T0 "
            sQuery = sQuery & " INNER JOIN [" & p_oCompDef.p_sDataBaseName & "].dbo.OCRC T1 ON T1.CreditCard = T0.CreditCard "
            sQuery = sQuery & " WHERE T1.UpdateDate >= DATEADD(DAY,DATEDIFF(DAY,0,GETDATE()),- " & p_oCompDef.p_iIntegDays & ")"

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Tender update Query exec " & sQuery, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)

            If ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc) <> RTN_SUCCESS Then
                Console.WriteLine("Error while updating Tender", sFuncName)
            Else
                Console.WriteLine("Tender update Successful", sFuncName)
            End If
            '********************************************************************************

            'Console.WriteLine("Stock Checking Query :", sFuncName)
            'sQuery = "[AE_SP002_GetNoStockItem]'[" & p_oCompDef.p_sDataBaseName & "]'"
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Stock Checking Query : " & sQuery, sFuncName)
            'If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
            'ExecuteSQLQuery_DT(P_sConString, sQuery, sErrDesc)

            Console.WriteLine("Completed with SUCCESS ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            Console.WriteLine("Completed with ERROR : " & sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)

        End Try

    End Sub

End Module
