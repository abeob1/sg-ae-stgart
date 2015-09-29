Module modDownpayment

    Public Function AR_DownPaymentInvoice(ByVal oDVARInvoice As DataView, ByVal oDVPayment As DataView, ByRef oCompany As SAPbobsCOM.Company, _
                                    ByVal sSODocEntry As String, ByRef sDocEntry As String, ByRef sDocNum As String, ByVal sCardCode As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim oARDownPayment As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
        Dim oARDownPayment_Doc As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
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
            sFuncName = "AR_DownPaymentInvoice()"
            Console.WriteLine("Starting Function ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function ", sFuncName)

            sQuery = "select T0.ItemCode , T0.Quantity , T0.LineTotal , T0.VatGroup, T0.WhsCode  from RDR1 T0 where t0.DocEntry = '" & sSODocEntry & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Getting Sales Order Information " & sQuery, sFuncName)
            oRset.DoQuery(sQuery)

            dIncomeDate = Convert.ToDateTime(oDVARInvoice.Item(0).Row("PHOSTxDate").ToString.Trim)
            sPOSNumber = CStr(oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim)
            sWhsCode = CStr(oDVARInvoice.Item(0).Row("HOutlet").ToString.Trim)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("AR Invoice dPostxdatetime " & dPostxdatetime, sFuncName)
            dPostxdatetime = oDVARInvoice.Item(0).Row("HPOSTxDatetime").ToString.Trim

            oARDownPayment.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDownPayments

            oARDownPayment.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice

            tDocTime = tDocTime.AddHours(0)
            tDocTime = tDocTime.AddMinutes(0)

            oARDownPayment.CardCode = p_oCompDef.p_sCardCode
            oARDownPayment.DocDate = dIncomeDate
            oARDownPayment.DocDueDate = dIncomeDate
            oARDownPayment.TaxDate = dIncomeDate
            oARDownPayment.NumAtCard = sWhsCode & " - " & sPOSNumber

            oARDownPayment.UserFields.Fields.Item("U_POS_RefNo").Value = oDVARInvoice.Item(0).Row("HPOSTxNo").ToString.Trim
            oARDownPayment.UserFields.Fields.Item("U_Date").Value = dIncomeDate
            oARDownPayment.UserFields.Fields.Item("U_Time").Value = dPostxdatetime

            ''For Each dvr As DataRowView In oDVARInvoice
            ''    oARDownPayment.Lines.ItemCode = dvr("DItemCode").ToString.Trim
            ''    oARDownPayment.Lines.Quantity = CDbl(dvr("DQuantity").ToString.Trim)
            ''    '' MsgBox(dvr("DPrice").ToString.Trim)
            ''    '' oARDownPayment.Lines.Price = CDbl(dvr("DPrice").ToString.Trim)
            ''    oARDownPayment.Lines.LineTotal = CDbl(dvr("DLineTotal").ToString.Trim)
            ''    oARDownPayment.Lines.WarehouseCode = sWhsCode
            ''    oARDownPayment.Lines.VatGroup = p_oCompDef.p_sZeroTax  'dvr("VatGourpSa").ToString.Trim
            ''    oARDownPayment.Lines.Add()
            ''Next

            For imjs As Integer = 1 To oRset.RecordCount
                oARDownPayment.Lines.ItemCode = oRset.Fields.Item("ItemCode").Value
                oARDownPayment.Lines.Quantity = oRset.Fields.Item("Quantity").Value
                oARDownPayment.Lines.LineTotal = oRset.Fields.Item("LineTotal").Value
                oARDownPayment.Lines.WarehouseCode = oRset.Fields.Item("WhsCode").Value
                oARDownPayment.Lines.VatGroup = p_oCompDef.p_sZeroTax  'dvr("VatGourpSa").ToString.Trim
                oARDownPayment.Lines.BaseType = "17"
                oARDownPayment.Lines.BaseEntry = sSODocEntry
                oARDownPayment.Lines.Add()
                oRset.MoveNext()
            Next imjs

            If oCompany.InTransaction = False Then oCompany.StartTransaction()
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add Draft ", sFuncName)
            lRetCode = oARDownPayment.Add()

            If lRetCode <> 0 Then
                sErrDesc = oCompany.GetLastErrorDescription
                Call WriteToLogFile(sErrDesc, sFuncName)
                ' oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
                ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARDownPayment)
                Return RTN_ERROR

            Else
                '' System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
                ''oARInvoice = Nothing
                '----------------- AR Invoice Draft Created Successfully
                oCompany.GetNewObjectCode(sDocEntry)
                Console.WriteLine("Draft Added Successfully " & sDocEntry, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Draft Added Successfully  " & sDocEntry, sFuncName)
                Console.WriteLine("Assigning Batch   " & sDocEntry, sFuncName)

                If oARDownPayment.GetByKey(sDocEntry) Then
                    dDocTotal = oARDownPayment.DocTotal
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

                                oARDownPayment.Lines.SetCurrentLine(irow)
                                oARDownPayment.Lines.BatchNumbers.SetCurrentLine(iloop)

                                If dInvQuantity > 0 Then
                                    If oDT_Batch.Rows.Count = 0 Then
                                        oARDownPayment.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                        If dInvQuantity > dBatchQuantity Then
                                            'If Balance Qty>Batch Qty, then get full Batch Qty
                                            oARDownPayment.Lines.BatchNumbers.Quantity = dBatchQuantity
                                            'minus current qty with Batch Qty
                                            dInvQuantity = dInvQuantity - dBatchQuantity
                                            oDT_Batch.Rows.Add(sProductCode, dBatchNumber, 0)
                                        Else
                                            oARDownPayment.Lines.BatchNumbers.Quantity = dInvQuantity
                                            oDT_Batch.Rows.Add(sProductCode, dBatchNumber, dBatchQuantity - dInvQuantity)
                                            dInvQuantity = dInvQuantity - dInvQuantity
                                        End If

                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                                        oARDownPayment.Lines.BatchNumbers.Add()
                                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)

                                        If dInvQuantity <= 0 Then Exit For
                                    Else
                                        oDV_Batch = New DataView(oDT_Batch)
                                        oDV_Batch.RowFilter = "ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'"
                                        If oDV_Batch.Count > 0 Then
                                            dRemBatchQuantity = oDV_Batch.Item(0).Row("Quantity")
                                            If dRemBatchQuantity > dInvQuantity Then
                                                oARDownPayment.Lines.BatchNumbers.Quantity = dInvQuantity
                                                oARDownPayment.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                                oRow = oDT_Batch.Select("ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'")
                                                oRow(0)("Quantity") = oDV_Batch.Item(0).Row("Quantity") - dInvQuantity
                                                oARDownPayment.Lines.BatchNumbers.Add()
                                                Exit For
                                            End If
                                            oARDownPayment.Lines.BatchNumbers.Quantity = dRemBatchQuantity
                                            oARDownPayment.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                            oRow = oDT_Batch.Select("ItemCode = '" & sProductCode & "' and BatchNum = '" & dBatchNumber & "'")
                                            oRow(0)("Quantity") = 0
                                            dInvQuantity = dInvQuantity - dRemBatchQuantity
                                            oARDownPayment.Lines.BatchNumbers.Add()

                                        Else
                                            oARDownPayment.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                            If dInvQuantity > dBatchQuantity Then
                                                'If Balance Qty>Batch Qty, then get full Batch Qty
                                                oARDownPayment.Lines.BatchNumbers.Quantity = dBatchQuantity
                                                'minus current qty with Batch Qty
                                                dInvQuantity = dInvQuantity - dBatchQuantity
                                                oDT_Batch.Rows.Add(sProductCode, dBatchNumber, 0)
                                            Else
                                                oARDownPayment.Lines.BatchNumbers.Quantity = dInvQuantity
                                                oDT_Batch.Rows.Add(sProductCode, dBatchNumber, dBatchQuantity - dInvQuantity)
                                                dInvQuantity = dInvQuantity - dInvQuantity
                                            End If

                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add BatchNumbers ", sFuncName)
                                            oARDownPayment.Lines.BatchNumbers.Add()
                                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Success - Add BatchNumbers ", sFuncName)
                                            If dInvQuantity <= 0 Then Exit For
                                        End If
                                    End If
                                Else
                                    '-------------------------- -ve quantity
                                    oARDownPayment.Lines.BatchNumbers.BatchNumber = dBatchNumber
                                    oARDownPayment.Lines.BatchNumbers.Quantity = Math.Abs(dInvQuantity)
                                    oARDownPayment.Lines.BatchNumbers.Add()
                                    Exit For
                                End If
                                oRset_Batch.MoveNext()
                            Next iloop
                        End If
                        oRset.MoveNext()
                    Next imjs

                    Dim dblRoundAmt As Double = 0.0
                    If oDVPayment.Count > 0 Then
                        dblRoundAmt = dPayamount - oARDownPayment.DocTotal
                    End If
                    If dblRoundAmt <> 0 Then
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calculating Rounding Amount: " & dblRoundAmt, sFuncName)
                        oARDownPayment.Lines.Add()
                        oARDownPayment.Lines.ItemCode = p_oCompDef.p_sRoundingItem
                        If dblRoundAmt > 0 Then
                            oARDownPayment.Lines.Quantity = 1
                        Else
                            oARDownPayment.Lines.Quantity = -1
                        End If
                        oARDownPayment.Lines.Price = Math.Abs(dblRoundAmt)
                        oARDownPayment.Lines.VatGroup = p_oCompDef.p_sZeroTax
                    End If

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting Update the AR Invoice Draft with Batch Information", sFuncName)
                    lRetCode = oARDownPayment.Update() 'Update the batch information
                    Console.WriteLine("Batch Updated Successfully " & sDocEntry, sFuncName)
                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        ' oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Update AR Invoice Draft) ", sFuncName)
                        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARDownPayment)
                        '' Return RTN_ERROR
                    End If

                    SARDraft = sDocEntry
                    Console.WriteLine("Attempting to Convert as a AR Invoice Document", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS (Update AR Invoice Draft) ", sFuncName)

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Convert as a AR Invoice Document ", sFuncName)

                    lRetCode = oARDownPayment.SaveDraftToDocument()

                    If lRetCode <> 0 Then
                        sErrDesc = oCompany.GetLastErrorDescription
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        ' oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "FAIL", sErrDesc, "", Now.ToShortTimeString, "", "")
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR (Convert as a AR Invoice Document) ", sFuncName)
                        ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARDownPayment)
                        '' Return RTN_ERROR
                    End If
                    oCompany.GetNewObjectCode(sDocEntry)
                    oARDownPayment_Doc.GetByKey(sDocEntry)
                    sDocNum = oARDownPayment_Doc.DocNum

                    Console.WriteLine("Converted To AR Invoice Successful " & sDocEntry, sFuncName)
                    '  oDTStatus.Rows.Add(oDVARInvoice.Item(0).Row("HTransID").ToString.Trim, "", "SUCCESS", "", "", Now.ToShortTimeString, sDocEntry, oARDownPayment_Doc.DocNum)
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
            oARDownPayment = Nothing
            Return RTN_ERROR
        Finally

            If oARDownPayment.GetByKey(SARDraft) Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Remove te Draft ", sFuncName)
                lRetCode = oARDownPayment.Remove()
                If lRetCode <> 0 Then
                    sErrDesc = oCompany.GetLastErrorDescription
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR ", sFuncName)
                    ''System.Runtime.InteropServices.Marshal.ReleaseComObject(oARDownPayment)
                End If
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARDownPayment)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARDownPayment_Doc)
        End Try
    End Function

End Module
