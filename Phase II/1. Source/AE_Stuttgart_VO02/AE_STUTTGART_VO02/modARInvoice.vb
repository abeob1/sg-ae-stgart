Module modARInvoice


    Public Function AR_InvoiceCreation(ByVal oDVARInvoice As DataView, ByVal oDVPayment As DataView, ByRef oCompany As SAPbobsCOM.Company, _
                                       ByRef sDocEntry As String, ByRef sDocNum As String, ByVal sCardCode As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oARInvoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
        Dim oARInvoice_Doc As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        Dim dIncomeDate As Date
        Dim tDocTime As DateTime
        Dim sWhsCode As String = String.Empty
        Dim sPOSNumber As String = String.Empty
        Dim sProductCode As String = String.Empty
        Dim sBOMCode As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim sQueryup As String = String.Empty
        Dim sManBatchItem As String = String.Empty
        Dim oBatchDT As DataTable = Nothing
        Dim dBatchQuantity As Double = 0
        Dim dRemBatchQuantity As Double = 0
        Dim dBatchNumber As String = String.Empty
        Dim dInvQuantity As Double
        Dim lRetCode As Integer
        Dim irow As Integer = 0
        Dim dDocTotal As Double = 0.0
        Dim oDV_BOM As DataView = New DataView(oDT_BOM)
        Dim oDT_Batch As DataTable = New DataTable
        Dim oDV_Batch As DataView = Nothing
        Dim oRow() As Data.DataRow = Nothing
        Dim SARDraft As String = String.Empty
        Dim dPostxdatetime As Date
        Dim oDT_Payamount As DataTable = New DataTable
        Dim dPayamount As Double
        oDT_Payamount = oDVPayment.ToTable

        If oDT_Payamount.Rows.Count > 0 Then
            dPayamount = Convert.ToDecimal(oDT_Payamount.Compute("sum(PaymentAmount)", String.Empty).ToString)
        End If

        oDT_Batch.Columns.Add("ItemCode", GetType(String))
        oDT_Batch.Columns.Add("BatchNum", GetType(String))
        oDT_Batch.Columns.Add("Quantity", GetType(Decimal))
        '' oDT_Batch.Columns.Add("date", GetType(Date))


        Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRset_Batch As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            sFuncName = "AR_InvoiceCreation()"
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            dIncomeDate = Convert.ToDateTime(oDVARInvoice.Item(0).Row("PHOSTxDate").ToString.Trim)
            sPOSNumber = CStr(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim)
            sWhsCode = CStr(oDVARInvoice.Item(0).Row("HOutlet").ToString.Trim)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AR Invoice dPostxdatetime " & dPostxdatetime, sFuncName)
            dPostxdatetime = oDVARInvoice.Item(0).Row("HPOSTxDatetime").ToString.Trim

            oARInvoice.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices

            tDocTime = tDocTime.AddHours(0)
            tDocTime = tDocTime.AddMinutes(0)

            oARInvoice.CardCode = p_oCompDef.p_sCardCode
            oARInvoice.DocDate = dIncomeDate
            oARInvoice.DocDueDate = dIncomeDate
            oARInvoice.TaxDate = dIncomeDate
            oARInvoice.NumAtCard = sWhsCode & " - " & sPOSNumber

            oARInvoice.UserFields.Fields.Item("U_POS_RefNo").Value = oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim
            oARInvoice.UserFields.Fields.Item("U_Date").Value = dIncomeDate
            oARInvoice.UserFields.Fields.Item("U_Time").Value = dPostxdatetime

            For Each dvr As DataRowView In oDVARInvoice
                oARInvoice.Lines.ItemCode = dvr("DItemCode").ToString.Trim
                oARInvoice.Lines.Quantity = CDbl(dvr("DQuantity").ToString.Trim)
                '' MsgBox(dvr("DPrice").ToString.Trim)
                '' oARInvoice.Lines.Price = CDbl(dvr("DPrice").ToString.Trim)
                oARInvoice.Lines.LineTotal = CDbl(dvr("DLineTotal").ToString.Trim)
                oARInvoice.Lines.WarehouseCode = sWhsCode
                oARInvoice.Lines.VatGroup = dvr("VatGourpSa").ToString.Trim
                oARInvoice.Lines.Add()
            Next

            If oCompany.InTransaction = False Then oCompany.StartTransaction()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add Draft ", sFuncName)
            lRetCode = oARInvoice.Add()

            If lRetCode <> 0 Then
                sErrDesc = oCompany.GetLastErrorDescription
                Call WriteToLogFile(sErrDesc, sFuncName)
                ' oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
                ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                Return RTN_ERROR

            Else
                '' System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                ''oARInvoice = Nothing
                '----------------- AR Invoice Draft Created Successfully
                oCompany.GetNewObjectCode(sDocEntry)
                Console.WriteLine("Draft Added Successfully " & sDocEntry, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Draft Added Successfully  " & sDocEntry, sFuncName)
                Console.WriteLine("Assigning Batch   " & sDocEntry, sFuncName)

                If oARInvoice.GetByKey(sDocEntry) Then
                    dDocTotal = oARInvoice.DocTotal
                    sQuery = "SELECT T0.[LineNum], T0.[ItemCode], T0.[Quantity], T0.[WhsCode], T1.[ManBtchNum] FROM DRF1 T0  INNER JOIN OITM T1 ON T0.[ItemCode] = T1.[ItemCode] WHERE T0.[DocEntry] = '" & sDocEntry & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Draft Details SQL " & sQuery, sFuncName)
                    oRset.DoQuery(sQuery)
                    For imjs As Integer = 0 To oRset.RecordCount - 1
                        sProductCode = oRset.Fields.Item("ItemCode").Value
                        sManBatchItem = oRset.Fields.Item("ManBtchNum").Value
                        irow = oRset.Fields.Item("LineNum").Value 'Row Number
                        dInvQuantity = CDbl(oRset.Fields.Item("Quantity").Value) 'Item Quantity
                        If sManBatchItem = "Y" Then
                            sQuery = "SELECT BatchNum ,Quantity , SysNumber  FROM OIBT WITH (NOLOCK) WHERE ItemCode ='" & sProductCode & "' and Quantity >0 " & _
                                              "AND WhsCode ='" & sWhsCode & "' ORDER BY InDate ASC "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Informations SQL " & sQuery, sFuncName)
                            oRset_Batch.DoQuery(sQuery)
                            For iloop As Integer = 0 To oRset_Batch.RecordCount - 1

                                dBatchQuantity = CDbl(oRset_Batch.Fields.Item("Quantity").Value) 'Batch Quantity
                                dBatchNumber = oRset_Batch.Fields.Item("BatchNum").Value 'Batch

                                oARInvoice.Lines.SetCurrentLine(irow)
                                oARInvoice.Lines.BatchNumbers.SetCurrentLine(iloop)

                                If dInvQuantity > 0 Then
                                    If oDT_Batch.Rows.Count = 0 Then
                                        oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                        If dInvQuantity > dBatchQuantity Then
                                            'If Balance Qty>Batch Qty, then get full Batch Qty
                                            oARInvoice.Lines.BatchNumbers.Quantity = dBatchQuantity
                                            'minus current qty with Batch Qty
                                            dInvQuantity = dInvQuantity - dBatchQuantity
                                            oDT_Batch.Rows.Add(sProductCode, dBatchNumber, 0)
                                        Else
                                            oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                            oDT_Batch.Rows.Add(sProductCode, dBatchNumber, dBatchQuantity - dInvQuantity)
                                            dInvQuantity = dInvQuantity - dInvQuantity
                                        End If

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                                        oARInvoice.Lines.BatchNumbers.Add()
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)

                                        If dInvQuantity <= 0 Then Exit For
                                    Else
                                        oDV_Batch = New DataView(oDT_Batch)
                                        oDV_Batch.RowFilter = "ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'"
                                        If oDV_Batch.Count > 0 Then
                                            dRemBatchQuantity = oDV_Batch.Item(0).Row("Quantity")
                                            If dRemBatchQuantity > dInvQuantity Then
                                                oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                                oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                                oRow = oDT_Batch.Select("ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'")
                                                oRow(0)("Quantity") = oDV_Batch.Item(0).Row("Quantity") - dInvQuantity
                                                oARInvoice.Lines.BatchNumbers.Add()
                                                Exit For
                                            End If
                                            oARInvoice.Lines.BatchNumbers.Quantity = dRemBatchQuantity
                                            oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                            oRow = oDT_Batch.Select("ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'")
                                            oRow(0)("Quantity") = 0
                                            dInvQuantity = dInvQuantity - dRemBatchQuantity
                                            oARInvoice.Lines.BatchNumbers.Add()

                                        Else
                                            oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                            If dInvQuantity > dBatchQuantity Then
                                                'If Balance Qty>Batch Qty, then get full Batch Qty
                                                oARInvoice.Lines.BatchNumbers.Quantity = dBatchQuantity
                                                'minus current qty with Batch Qty
                                                dInvQuantity = dInvQuantity - dBatchQuantity
                                                oDT_Batch.Rows.Add(sProductCode, dBatchNumber, 0)
                                            Else
                                                oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                                oDT_Batch.Rows.Add(sProductCode, dBatchNumber, dBatchQuantity - dInvQuantity)
                                                dInvQuantity = dInvQuantity - dInvQuantity
                                            End If

                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                                            oARInvoice.Lines.BatchNumbers.Add()
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)
                                            If dInvQuantity <= 0 Then Exit For
                                        End If
                                    End If
                                Else
                                    '-------------------------- -ve quantity
                                    oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                    oARInvoice.Lines.BatchNumbers.Quantity = Math.Abs(dInvQuantity)
                                    oARInvoice.Lines.BatchNumbers.Add()
                                    Exit For
                                End If
                                oRset_Batch.MoveNext()
                            Next iloop
                        End If
                        oRset.MoveNext()
                    Next imjs

                    Dim dblRoundAmt As Double = 0.0
                    If oDVPayment.Count > 0 Then
                        dblRoundAmt = dPayamount - oARInvoice.DocTotal
                    End If
                    If dblRoundAmt <> 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calculating Rounding Amount: " & dblRoundAmt, sFuncName)
                        oARInvoice.Lines.Add()
                        oARInvoice.Lines.ItemCode = p_oCompDef.p_sRoundingItem
                        If dblRoundAmt > 0 Then
                            oARInvoice.Lines.Quantity = 1
                        Else
                            oARInvoice.Lines.Quantity = -1
                        End If
                        oARInvoice.Lines.Price = Math.Abs(dblRoundAmt)
                        oARInvoice.Lines.VatGroup = p_oCompDef.p_sZeroTax
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Update the AR Invoice Draft with Batch Information", sFuncName)
                    lRetCode = oARInvoice.Update() 'Update the batch information
                    Console.WriteLine("Batch Updated Successfully " & sDocEntry, sFuncName)
                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        ' oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Update AR Invoice Draft) ", sFuncName)
                        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                        '' Return RTN_ERROR
                    End If

                    SARDraft = sDocEntry
                    Console.WriteLine("Attempting to Convert as a AR Invoice Document", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS (Update AR Invoice Draft) ", sFuncName)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Convert as a AR Invoice Document ", sFuncName)

                    lRetCode = oARInvoice.SaveDraftToDocument()

                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        ' oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Convert as a AR Invoice Document) ", sFuncName)
                        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                        '' Return RTN_ERROR
                    End If
                    oCompany.GetNewObjectCode(sDocEntry)
                    oARInvoice_Doc.GetByKey(sDocEntry)
                    sDocNum = oARInvoice_Doc.DocEntry

                    Console.WriteLine("Converted To AR Invoice Successful " & sDocEntry, sFuncName)
                    '  oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sDocEntry, oARInvoice_Doc.DocNum)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS (Convert as a AR Invoice Document) " & sDocEntry, sFuncName)
                End If

                Return RTN_SUCCESS
            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Console.WriteLine("Completed with ERROR", sFuncName)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            oARInvoice = Nothing
            Return RTN_ERROR
        Finally

            If oARInvoice.GetByKey(SARDraft) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Remove te Draft ", sFuncName)
                lRetCode = oARInvoice.Remove()
                If lRetCode <> 0 Then
                    sErrDesc = oCompany.GetLastErrorDescription
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
                    ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice_Doc)
        End Try
    End Function
    Public Function AR_InvoiceCreationSO(ByVal oDVARInvoice As DataView, ByVal oDVPayment As DataView, ByRef oCompany As SAPbobsCOM.Company, _
                                     ByVal sSODocEntry As String, ByRef sDocEntry As String, ByRef sDocNum As String, ByVal sCardCode As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oARInvoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
        Dim oARInvoice_Doc As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        Dim dIncomeDate As Date
        Dim tDocTime As DateTime
        Dim sWhsCode As String = String.Empty
        Dim sPOSNumber As String = String.Empty
        Dim sProductCode As String = String.Empty
        Dim sBOMCode As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim sQueryup As String = String.Empty
        Dim sManBatchItem As String = String.Empty
        Dim oBatchDT As DataTable = Nothing
        Dim dBatchQuantity As Double = 0
        Dim dRemBatchQuantity As Double = 0
        Dim dBatchNumber As String = String.Empty
        Dim dInvQuantity As Double
        Dim lRetCode As Integer
        Dim irow As Integer = 0
        Dim dDocTotal As Double = 0.0
        Dim oDV_BOM As DataView = New DataView(oDT_BOM)
        Dim oDT_Batch As DataTable = New DataTable
        Dim oDV_Batch As DataView = Nothing
        Dim oRow() As Data.DataRow = Nothing
        Dim SARDraft As String = String.Empty
        Dim dPostxdatetime As Date
        Dim oDT_Payamount As DataTable = New DataTable
        Dim dPayamount As Double
        oDT_Payamount = oDVPayment.ToTable

        If oDT_Payamount.Rows.Count > 0 Then
            dPayamount = Convert.ToDecimal(oDT_Payamount.Compute("sum(PaymentAmount)", String.Empty).ToString)
        End If

        oDT_Batch.Columns.Add("ItemCode", GetType(String))
        oDT_Batch.Columns.Add("BatchNum", GetType(String))
        oDT_Batch.Columns.Add("Quantity", GetType(Decimal))
        '' oDT_Batch.Columns.Add("date", GetType(Date))


        Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRset_Batch As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            sFuncName = "AR_InvoiceCreation()"
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)
            sQuery = "select T0.ItemCode , T0.Quantity , T0.LineTotal , T0.VatGroup, T0.WhsCode,T0.LineNum  from RDR1 T0 where t0.DocEntry = '" & sSODocEntry & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Sales Order Information " & sQuery, sFuncName)
            oRset.DoQuery(sQuery)

            dIncomeDate = Convert.ToDateTime(oDVARInvoice.Item(0).Row("PHOSTxDate").ToString.Trim)
            sPOSNumber = CStr(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim)
            sWhsCode = CStr(oDVARInvoice.Item(0).Row("HOutlet").ToString.Trim)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AR Invoice dPostxdatetime " & dPostxdatetime, sFuncName)
            dPostxdatetime = oDVARInvoice.Item(0).Row("HPOSTxDatetime").ToString.Trim

            oARInvoice.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices

            tDocTime = tDocTime.AddHours(0)
            tDocTime = tDocTime.AddMinutes(0)

            oARInvoice.CardCode = p_oCompDef.p_sCardCode
            oARInvoice.DocDate = dIncomeDate
            oARInvoice.DocDueDate = dIncomeDate
            oARInvoice.TaxDate = dIncomeDate
            oARInvoice.NumAtCard = sWhsCode & " - " & sPOSNumber

            oARInvoice.UserFields.Fields.Item("U_POS_RefNo").Value = oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim
            oARInvoice.UserFields.Fields.Item("U_Date").Value = dIncomeDate
            oARInvoice.UserFields.Fields.Item("U_Time").Value = dPostxdatetime

            For imjs As Integer = 1 To oRset.RecordCount
                oARInvoice.Lines.ItemCode = oRset.Fields.Item("ItemCode").Value
                oARInvoice.Lines.Quantity = oRset.Fields.Item("Quantity").Value
                oARInvoice.Lines.LineTotal = oRset.Fields.Item("LineTotal").Value
                oARInvoice.Lines.WarehouseCode = oRset.Fields.Item("WhsCode").Value
                oARInvoice.Lines.VatGroup = oRset.Fields.Item("VatGroup").Value
                oARInvoice.Lines.BaseType = "17"
                oARInvoice.Lines.BaseEntry = sSODocEntry
                oARInvoice.Lines.BaseLine = oRset.Fields.Item("LineNum").Value
                oARInvoice.Lines.Add()
                oRset.MoveNext()
            Next imjs

            ''  If oCompany.InTransaction = False Then oCompany.StartTransaction()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add Draft ", sFuncName)
            lRetCode = oARInvoice.Add()

            If lRetCode <> 0 Then
                sErrDesc = oCompany.GetLastErrorDescription
                Call WriteToLogFile(sErrDesc, sFuncName)
                ' oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
                ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                Return RTN_ERROR

            Else
                '' System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                ''oARInvoice = Nothing
                '----------------- AR Invoice Draft Created Successfully
                oCompany.GetNewObjectCode(sDocEntry)
                Console.WriteLine("Draft Added Successfully " & sDocEntry, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Draft Added Successfully  " & sDocEntry, sFuncName)
                Console.WriteLine("Assigning Batch   " & sDocEntry, sFuncName)

                If oARInvoice.GetByKey(sDocEntry) Then
                    dDocTotal = oARInvoice.DocTotal
                    sQuery = "SELECT T0.[LineNum], T0.[ItemCode], T0.[Quantity], T0.[WhsCode], T1.[ManBtchNum] FROM DRF1 T0  INNER JOIN OITM T1 ON T0.[ItemCode] = T1.[ItemCode] WHERE T0.[DocEntry] = '" & sDocEntry & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Draft Details SQL " & sQuery, sFuncName)
                    oRset.DoQuery(sQuery)
                    For imjs As Integer = 0 To oRset.RecordCount - 1
                        sProductCode = oRset.Fields.Item("ItemCode").Value
                        sManBatchItem = oRset.Fields.Item("ManBtchNum").Value
                        irow = oRset.Fields.Item("LineNum").Value 'Row Number
                        dInvQuantity = CDbl(oRset.Fields.Item("Quantity").Value) 'Item Quantity
                        If sManBatchItem = "Y" Then
                            sQuery = "SELECT BatchNum ,Quantity , SysNumber  FROM OIBT WITH (NOLOCK) WHERE ItemCode ='" & sProductCode & "' and Quantity >0 " & _
                                              "AND WhsCode ='" & sWhsCode & "' ORDER BY InDate ASC "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Informations SQL " & sQuery, sFuncName)
                            oRset_Batch.DoQuery(sQuery)
                            For iloop As Integer = 0 To oRset_Batch.RecordCount - 1

                                dBatchQuantity = CDbl(oRset_Batch.Fields.Item("Quantity").Value) 'Batch Quantity
                                dBatchNumber = oRset_Batch.Fields.Item("BatchNum").Value 'Batch

                                oARInvoice.Lines.SetCurrentLine(irow)
                                oARInvoice.Lines.BatchNumbers.SetCurrentLine(iloop)

                                If dInvQuantity > 0 Then
                                    If oDT_Batch.Rows.Count = 0 Then
                                        oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                        If dInvQuantity > dBatchQuantity Then
                                            'If Balance Qty>Batch Qty, then get full Batch Qty
                                            oARInvoice.Lines.BatchNumbers.Quantity = dBatchQuantity
                                            'minus current qty with Batch Qty
                                            dInvQuantity = dInvQuantity - dBatchQuantity
                                            oDT_Batch.Rows.Add(sProductCode, dBatchNumber, 0)
                                        Else
                                            oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                            oDT_Batch.Rows.Add(sProductCode, dBatchNumber, dBatchQuantity - dInvQuantity)
                                            dInvQuantity = dInvQuantity - dInvQuantity
                                        End If

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                                        oARInvoice.Lines.BatchNumbers.Add()
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)

                                        If dInvQuantity <= 0 Then Exit For
                                    Else
                                        oDV_Batch = New DataView(oDT_Batch)
                                        oDV_Batch.RowFilter = "ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'"
                                        If oDV_Batch.Count > 0 Then
                                            dRemBatchQuantity = oDV_Batch.Item(0).Row("Quantity")
                                            If dRemBatchQuantity > dInvQuantity Then
                                                oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                                oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                                oRow = oDT_Batch.Select("ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'")
                                                oRow(0)("Quantity") = oDV_Batch.Item(0).Row("Quantity") - dInvQuantity
                                                oARInvoice.Lines.BatchNumbers.Add()
                                                Exit For
                                            End If
                                            oARInvoice.Lines.BatchNumbers.Quantity = dRemBatchQuantity
                                            oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                            oRow = oDT_Batch.Select("ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'")
                                            oRow(0)("Quantity") = 0
                                            dInvQuantity = dInvQuantity - dRemBatchQuantity
                                            oARInvoice.Lines.BatchNumbers.Add()

                                        Else
                                            oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                            If dInvQuantity > dBatchQuantity Then
                                                'If Balance Qty>Batch Qty, then get full Batch Qty
                                                oARInvoice.Lines.BatchNumbers.Quantity = dBatchQuantity
                                                'minus current qty with Batch Qty
                                                dInvQuantity = dInvQuantity - dBatchQuantity
                                                oDT_Batch.Rows.Add(sProductCode, dBatchNumber, 0)
                                            Else
                                                oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                                oDT_Batch.Rows.Add(sProductCode, dBatchNumber, dBatchQuantity - dInvQuantity)
                                                dInvQuantity = dInvQuantity - dInvQuantity
                                            End If

                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                                            oARInvoice.Lines.BatchNumbers.Add()
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)
                                            If dInvQuantity <= 0 Then Exit For
                                        End If
                                    End If
                                Else
                                    '-------------------------- -ve quantity
                                    oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                    oARInvoice.Lines.BatchNumbers.Quantity = Math.Abs(dInvQuantity)
                                    oARInvoice.Lines.BatchNumbers.Add()
                                    Exit For
                                End If
                                oRset_Batch.MoveNext()
                            Next iloop
                        End If
                        oRset.MoveNext()
                    Next imjs

                    Dim dblRoundAmt As Double = 0.0
                    If oDVPayment.Count > 0 Then
                        dblRoundAmt = dPayamount - oARInvoice.DocTotal
                    End If
                    If dblRoundAmt <> 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calculating Rounding Amount: " & dblRoundAmt, sFuncName)
                        oARInvoice.Lines.Add()
                        oARInvoice.Lines.ItemCode = p_oCompDef.p_sRoundingItem
                        If dblRoundAmt > 0 Then
                            oARInvoice.Lines.Quantity = 1
                        Else
                            oARInvoice.Lines.Quantity = -1
                        End If
                        oARInvoice.Lines.Price = Math.Abs(dblRoundAmt)
                        oARInvoice.Lines.VatGroup = p_oCompDef.p_sZeroTax
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Update the AR Invoice Draft with Batch Information", sFuncName)
                    lRetCode = oARInvoice.Update() 'Update the batch information
                    Console.WriteLine("Batch Updated Successfully " & sDocEntry, sFuncName)
                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        ' oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Update AR Invoice Draft) ", sFuncName)
                        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                        '' Return RTN_ERROR
                    End If

                    SARDraft = sDocEntry
                    Console.WriteLine("Attempting to Convert as a AR Invoice Document", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS (Update AR Invoice Draft) ", sFuncName)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Convert as a AR Invoice Document ", sFuncName)

                    lRetCode = oARInvoice.SaveDraftToDocument()

                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        ' oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Convert as a AR Invoice Document) ", sFuncName)
                        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                        '' Return RTN_ERROR
                    End If
                    oCompany.GetNewObjectCode(sDocEntry)
                    oARInvoice_Doc.GetByKey(sDocEntry)
                    sDocNum = oARInvoice_Doc.DocNum

                    Console.WriteLine("Converted To AR Invoice Successful " & sDocEntry, sFuncName)
                    '  oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sDocEntry, oARInvoice_Doc.DocNum)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS (Convert as a AR Invoice Document) " & sDocEntry, sFuncName)
                End If

                Return RTN_SUCCESS
            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Console.WriteLine("Completed with ERROR", sFuncName)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            oARInvoice = Nothing
            Return RTN_ERROR
        Finally

            If oARInvoice.GetByKey(SARDraft) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Remove te Draft ", sFuncName)
                lRetCode = oARInvoice.Remove()
                If lRetCode <> 0 Then
                    sErrDesc = oCompany.GetLastErrorDescription
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
                    ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice_Doc)
        End Try
    End Function

    Public Function AR_InvoiceCreation_OLD(ByVal oDVARInvoice As DataView, ByVal oDVPayment As DataView, ByRef oCompany As SAPbobsCOM.Company, ByVal oDTStatus As DataTable, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oARInvoice As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
        Dim oARInvoice_Doc As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        Dim dIncomeDate As Date
        Dim tDocTime As DateTime
        Dim sWhsCode As String = String.Empty
        Dim sPOSNumber As String = String.Empty
        Dim sProductCode As String = String.Empty
        Dim sBOMCode As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim sQueryup As String = String.Empty
        Dim sManBatchItem As String = String.Empty
        Dim oBatchDT As DataTable = Nothing
        Dim dBatchQuantity As Double = 0
        Dim dRemBatchQuantity As Double = 0
        Dim dBatchNumber As String = String.Empty
        Dim dInvQuantity As Double
        Dim sDocEntry As String = String.Empty
        Dim lRetCode As Integer
        Dim irow As Integer = 0
        Dim dDocTotal As Double = 0.0
        oDTStatus.Clear()
        Dim oDV_BOM As DataView = New DataView(oDT_BOM)
        Dim oDT_Batch As DataTable = New DataTable
        Dim oDV_Batch As DataView = Nothing
        Dim oRow() As Data.DataRow = Nothing
        Dim SARDraft As String = String.Empty
        Dim dPostxdatetime As Date
        Dim oDT_Payamount As DataTable = New DataTable
        Dim dPayamount As Double
        oDT_Payamount = oDVPayment.ToTable

        If oDT_Payamount.Rows.Count > 0 Then
            dPayamount = Convert.ToDecimal(oDT_Payamount.Compute("sum(PaymentAmount)", String.Empty).ToString)
        End If

        oDT_Batch.Columns.Add("ItemCode", GetType(String))
        oDT_Batch.Columns.Add("BatchNum", GetType(String))
        oDT_Batch.Columns.Add("Quantity", GetType(Decimal))
        '' oDT_Batch.Columns.Add("date", GetType(Date))


        Dim oRset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim oRset_Batch As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Try
            sFuncName = "AR_InvoiceCreation()"
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            dIncomeDate = Convert.ToDateTime(oDVARInvoice.Item(0).Row("PHOSTxDate").ToString.Trim)
            sPOSNumber = CStr(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim)
            sWhsCode = CStr(oDVARInvoice.Item(0).Row("HOutlet").ToString.Trim)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AR Invoice dPostxdatetime " & dPostxdatetime, sFuncName)
            dPostxdatetime = oDVARInvoice.Item(0).Row("HPOSTxDatetime").ToString.Trim

            oARInvoice.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices

            tDocTime = tDocTime.AddHours(0)
            tDocTime = tDocTime.AddMinutes(0)

            oARInvoice.CardCode = p_oCompDef.p_sCardCode
            oARInvoice.DocDate = dIncomeDate
            oARInvoice.DocDueDate = dIncomeDate
            oARInvoice.TaxDate = dIncomeDate
            oARInvoice.NumAtCard = sWhsCode & " - " & sPOSNumber

            oARInvoice.UserFields.Fields.Item("U_POS_RefNo").Value = oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim
            oARInvoice.UserFields.Fields.Item("U_Date").Value = dIncomeDate
            oARInvoice.UserFields.Fields.Item("U_Time").Value = dPostxdatetime

            For Each dvr As DataRowView In oDVARInvoice
                oARInvoice.Lines.ItemCode = dvr("DItemCode").ToString.Trim
                oARInvoice.Lines.Quantity = CDbl(dvr("DQuantity").ToString.Trim)
                '' MsgBox(dvr("DPrice").ToString.Trim)
                '' oARInvoice.Lines.Price = CDbl(dvr("DPrice").ToString.Trim)
                oARInvoice.Lines.LineTotal = CDbl(dvr("DLineTotal").ToString.Trim)
                oARInvoice.Lines.WarehouseCode = sWhsCode
                oARInvoice.Lines.VatGroup = dvr("VatGourpSa").ToString.Trim
                oARInvoice.Lines.Add()
            Next

            If oCompany.InTransaction = False Then oCompany.StartTransaction()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add Draft ", sFuncName)
            lRetCode = oARInvoice.Add()

            If lRetCode <> 0 Then
                sErrDesc = oCompany.GetLastErrorDescription
                Call WriteToLogFile(sErrDesc, sFuncName)
                oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
                ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                Return RTN_ERROR

            Else
                '' System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                ''oARInvoice = Nothing
                '----------------- AR Invoice Draft Created Successfully
                oCompany.GetNewObjectCode(sDocEntry)
                Console.WriteLine("Draft Added Successfully " & sDocEntry, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Draft Added Successfully  " & sDocEntry, sFuncName)
                Console.WriteLine("Assigning Batch   " & sDocEntry, sFuncName)

                If oARInvoice.GetByKey(sDocEntry) Then
                    dDocTotal = oARInvoice.DocTotal
                    sQuery = "SELECT T0.[LineNum], T0.[ItemCode], T0.[Quantity], T0.[WhsCode], T1.[ManBtchNum] FROM DRF1 T0  INNER JOIN OITM T1 ON T0.[ItemCode] = T1.[ItemCode] WHERE T0.[DocEntry] = '" & sDocEntry & "'"
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Draft Details SQL " & sQuery, sFuncName)
                    oRset.DoQuery(sQuery)
                    For imjs As Integer = 0 To oRset.RecordCount - 1
                        sProductCode = oRset.Fields.Item("ItemCode").Value
                        sManBatchItem = oRset.Fields.Item("ManBtchNum").Value
                        irow = oRset.Fields.Item("LineNum").Value 'Row Number
                        dInvQuantity = CDbl(oRset.Fields.Item("Quantity").Value) 'Item Quantity
                        If sManBatchItem = "Y" Then
                            sQuery = "SELECT BatchNum ,Quantity , SysNumber  FROM OIBT WITH (NOLOCK) WHERE ItemCode ='" & sProductCode & "' and Quantity >0 " & _
                                              "AND WhsCode ='" & sWhsCode & "' ORDER BY InDate ASC "
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Batch Informations SQL " & sQuery, sFuncName)
                            oRset_Batch.DoQuery(sQuery)
                            For iloop As Integer = 0 To oRset_Batch.RecordCount - 1

                                dBatchQuantity = CDbl(oRset_Batch.Fields.Item("Quantity").Value) 'Batch Quantity
                                dBatchNumber = oRset_Batch.Fields.Item("BatchNum").Value 'Batch

                                oARInvoice.Lines.SetCurrentLine(irow)
                                oARInvoice.Lines.BatchNumbers.SetCurrentLine(iloop)

                                If dInvQuantity > 0 Then
                                    If oDT_Batch.Rows.Count = 0 Then
                                        oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                        If dInvQuantity > dBatchQuantity Then
                                            'If Balance Qty>Batch Qty, then get full Batch Qty
                                            oARInvoice.Lines.BatchNumbers.Quantity = dBatchQuantity
                                            'minus current qty with Batch Qty
                                            dInvQuantity = dInvQuantity - dBatchQuantity
                                            oDT_Batch.Rows.Add(sProductCode, dBatchNumber, 0)
                                        Else
                                            oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                            oDT_Batch.Rows.Add(sProductCode, dBatchNumber, dBatchQuantity - dInvQuantity)
                                            dInvQuantity = dInvQuantity - dInvQuantity
                                        End If

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                                        oARInvoice.Lines.BatchNumbers.Add()
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)

                                        If dInvQuantity <= 0 Then Exit For
                                    Else
                                        oDV_Batch = New DataView(oDT_Batch)
                                        oDV_Batch.RowFilter = "ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'"
                                        If oDV_Batch.Count > 0 Then
                                            dRemBatchQuantity = oDV_Batch.Item(0).Row("Quantity")
                                            If dRemBatchQuantity > dInvQuantity Then
                                                oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                                oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                                oRow = oDT_Batch.Select("ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'")
                                                oRow(0)("Quantity") = oDV_Batch.Item(0).Row("Quantity") - dInvQuantity
                                                oARInvoice.Lines.BatchNumbers.Add()
                                                Exit For
                                            End If
                                            oARInvoice.Lines.BatchNumbers.Quantity = dRemBatchQuantity
                                            oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                            oRow = oDT_Batch.Select("ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'")
                                            oRow(0)("Quantity") = 0
                                            dInvQuantity = dInvQuantity - dRemBatchQuantity
                                            oARInvoice.Lines.BatchNumbers.Add()

                                        Else
                                            oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                            If dInvQuantity > dBatchQuantity Then
                                                'If Balance Qty>Batch Qty, then get full Batch Qty
                                                oARInvoice.Lines.BatchNumbers.Quantity = dBatchQuantity
                                                'minus current qty with Batch Qty
                                                dInvQuantity = dInvQuantity - dBatchQuantity
                                                oDT_Batch.Rows.Add(sProductCode, dBatchNumber, 0)
                                            Else
                                                oARInvoice.Lines.BatchNumbers.Quantity = dInvQuantity
                                                oDT_Batch.Rows.Add(sProductCode, dBatchNumber, dBatchQuantity - dInvQuantity)
                                                dInvQuantity = dInvQuantity - dInvQuantity
                                            End If

                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                                            oARInvoice.Lines.BatchNumbers.Add()
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)
                                            If dInvQuantity <= 0 Then Exit For
                                        End If
                                    End If
                                Else
                                    '-------------------------- -ve quantity
                                    oARInvoice.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                    oARInvoice.Lines.BatchNumbers.Quantity = Math.Abs(dInvQuantity)
                                    oARInvoice.Lines.BatchNumbers.Add()
                                    Exit For
                                End If
                                oRset_Batch.MoveNext()
                            Next iloop
                        End If
                        oRset.MoveNext()
                    Next imjs

                    Dim dblRoundAmt As Double = 0.0
                    If oDVPayment.Count > 0 Then
                        dblRoundAmt = dPayamount - oARInvoice.DocTotal
                    End If
                    If dblRoundAmt <> 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calculating Rounding Amount: " & dblRoundAmt, sFuncName)
                        oARInvoice.Lines.Add()
                        oARInvoice.Lines.ItemCode = p_oCompDef.p_sRoundingItem
                        If dblRoundAmt > 0 Then
                            oARInvoice.Lines.Quantity = 1
                        Else
                            oARInvoice.Lines.Quantity = -1
                        End If
                        oARInvoice.Lines.Price = Math.Abs(dblRoundAmt)
                        oARInvoice.Lines.VatGroup = p_oCompDef.p_sZeroTax
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Update the AR Invoice Draft with Batch Information", sFuncName)
                    lRetCode = oARInvoice.Update() 'Update the batch information
                    Console.WriteLine("Batch Updated Successfully " & sDocEntry, sFuncName)
                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Update AR Invoice Draft) ", sFuncName)
                        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                        '' Return RTN_ERROR
                        GoTo ERRORDISPLAY
                    End If

                    SARDraft = sDocEntry
                    Console.WriteLine("Attempting to Convert as a AR Invoice Document", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS (Update AR Invoice Draft) ", sFuncName)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Convert as a AR Invoice Document ", sFuncName)

                    lRetCode = oARInvoice.SaveDraftToDocument()

                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Convert as a AR Invoice Document) ", sFuncName)
                        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                        '' Return RTN_ERROR
                        GoTo ERRORDISPLAY
                    End If
                    oCompany.GetNewObjectCode(sDocEntry)
                    oARInvoice_Doc.GetByKey(sDocEntry)

                    Console.WriteLine("Converted To AR Invoice Successful " & sDocEntry, sFuncName)
                    oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sDocEntry, oARInvoice_Doc.DocNum)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS (Convert as a AR Invoice Document) " & sDocEntry, sFuncName)
                End If


                If oDVARInvoice.Item(0).Row("HPOSTxType").ToString.Trim = "S" Then
                    '************************************ Incoming Payment Started ************************************************************************************



                    If oDVPayment Is Nothing Then
                        Console.WriteLine("No matching records found in Payement Table " & sDocEntry, sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No matching records found in Payement Table : AR Invoice DocEntry " & sDocEntry, sFuncName)
                    Else
                        If oDVPayment.Count > 0 Then
                            Console.WriteLine("Calling Funcion AR_IncomingPayment() " & sDocEntry, sFuncName)
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_IncomingPayment() : AR Invoice DocEntry " & sDocEntry, sFuncName)
                            If AR_IncomingPayment(oDVPayment, oCompany, sDocEntry, dIncomeDate, sPOSNumber _
                                                                     , sWhsCode, p_oCompDef.p_sCardCode, sErrDesc) <> RTN_SUCCESS Then

                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Console.WriteLine("Completed with ERROR", sFuncName)
                                oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString)
                                Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                                If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                oARInvoice = Nothing
                                ''  Return RTN_ERROR
                            End If
                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No matching records found in Payement Table : AR Invoice DocEntry " & sDocEntry, sFuncName)
                            Console.WriteLine("No matching records found in Payement Table " & sDocEntry, sFuncName)
                        End If
                    End If


                ElseIf oDVARInvoice.Item(0).Row("HPOSTxType").ToString.Trim = "V" Then
                    '************************************ AR Credit Memo ************************************************************************************
                    Console.WriteLine("Calling Funcion AR_CreditMemo() " & sDocEntry, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling Funcion AR_CreditMemo() : AR Invoice DocEntry " & sDocEntry, sFuncName)

                    If AR_CreditMemo(oCompany, sDocEntry, dIncomeDate, sErrDesc) <> RTN_SUCCESS Then

                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Console.WriteLine("Completed with ERROR", sFuncName)
                        oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString)
                        Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
                        If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        oARInvoice = Nothing
                        ''Return RTN_ERROR
                    End If

                End If
                sErrDesc = ""
                ''  oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString)

ERRORDISPLAY:   If oDTStatus Is Nothing Then
                Else
                    Dim sTrandID As String = String.Empty
                    Dim dSyncDatetime As DateTime
                    For imjs As Integer = 0 To oDTStatus.Rows.Count - 1

                        If sTrandID <> oDTStatus.Rows(imjs).Item("HID").ToString.Trim Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Date Time " & Now.Date & " " & oDTStatus.Rows(imjs).Item("Time").ToString.Trim, sFuncName)
                            dSyncDatetime = Now.Date & " " & oDTStatus.Rows(imjs).Item("Time").ToString.Trim
                            sQueryup += "UPDATE " & p_oCompDef.p_sIntDBName & ".. [AB_SalesTransHeader]" & _
    "SET [Status] = '" & oDTStatus.Rows(imjs).Item("Status").ToString.Trim & "' ,[ErrorMsg] = '" & oDTStatus.Rows(imjs).Item("HErrorMsg").ToString.Trim & "' , " & _
    "[SAPSyncDate] =  DATEADD(day,datediff(day,0,GETDATE()),0) ,[SAPSyncDateTime] = GETDATE(), [ARDocEntry] = '" & oDTStatus.Rows(imjs).Item("DocEntry").ToString.Trim & "' " & _
    "WHERE [ID] = '" & oDTStatus.Rows(imjs).Item("HID").ToString.Trim & "'"
                            sTrandID = oDTStatus.Rows(imjs).Item("HID").ToString.Trim

                            sQueryup += "UPDATE " & p_oCompDef.p_sIntDBName & ".. [AB_SalesTransDetail] SET [ErrMsg] = '' " & _
    " WHERE [HeaderID] = '" & oDTStatus.Rows(imjs).Item("HID").ToString.Trim & "'"
                        End If

                        If Not String.IsNullOrEmpty(oDTStatus.Rows(imjs).Item("LErrorMsg").ToString.Trim) Then
                            sQueryup += "UPDATE " & p_oCompDef.p_sIntDBName & ".. [AB_SalesTransDetail] SET [ErrMsg] = '" & oDTStatus.Rows(imjs).Item("LErrorMsg").ToString.Trim & "' " & _
        " WHERE [HeaderID] = '" & oDTStatus.Rows(imjs).Item("HID").ToString.Trim & "' and [ItemCode] = '" & oDTStatus.Rows(imjs).Item("LItem").ToString.Trim & "' "
                        End If

                    Next imjs
                    oDTStatus.Clear()
                    sTrandID = String.Empty

                End If

                If sQueryup.Length > 1 Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Validation Update SQL " & sQueryup, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Executing Query", sFuncName)
                    oRset.DoQuery(sQueryup)
                End If

                If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                Console.WriteLine("Committed the Transaction for TransID " & oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Committed the Transaction Reference POSNumber : " & sPOSNumber, sFuncName)
                ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                '' If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Disconnecting the Company and Release the Object ", sFuncName)

                Return RTN_SUCCESS
            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Console.WriteLine("Completed with ERROR", sFuncName)
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Console.WriteLine("Rollback the SAP Transaction ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Rollback the SAP Transaction ", sFuncName)
            If oCompany.InTransaction = True Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            oARInvoice = Nothing
            Return RTN_ERROR
        Finally

            If oARInvoice.GetByKey(SARDraft) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Remove te Draft ", sFuncName)
                lRetCode = oARInvoice.Remove()
                If lRetCode <> 0 Then
                    sErrDesc = oCompany.GetLastErrorDescription
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
                    ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice_Doc)
        End Try
    End Function

End Module
