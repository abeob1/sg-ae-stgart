Module modSalesOrder


    Function Salesorder(ByRef oDICompany As SAPbobsCOM.Company, _
                           ByVal sDocEntry As String, ByRef oDVARInvoice As DataView, _
                           ByVal sHeaderID As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim lRetCode As Long
        Dim oSalesOrder As SAPbobsCOM.Documents
        Dim sPayDocEntry As String = String.Empty
        Dim sQuery As String = String.Empty
        Dim oDT_Salesorder As New DataTable
        Dim oDT_INTItemDetails As New DataTable
        Dim oDT_INTItemDetails_Final As New DataTable
        Dim iLineNo As Integer = 0

        '' this is a change

        oSalesOrder = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
        Try
            sFuncName = "Salesorder()"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)
            sQuery = "SELECT T0.[LineNum], T0.[ItemCode] FROM [" & p_oCompDef.p_sDataBaseName & "].dbo.RDR1 T0 WHERE T0.[DocEntry] = '" & sDocEntry & "'"
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Fetching Sales Order Information : " & sQuery, sFuncName)

            If oSalesOrder.GetByKey(sDocEntry) Then
                Console.WriteLine("Calling ExecuteSQLQuery_DT() ", sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ExecuteSQLQuery_DT()", sFuncName)
                oDT_Salesorder = ExecuteSQLQuery_DT(P_sConString, sQuery)

                'oDT_INTItemDetails = oDVARInvoice.Table.DefaultView.ToTable(True, "DItemCode", "DQuantity", "DLineTotal", "DOutlet", "VatGourpSa")
                oDT_INTItemDetails = oDVARInvoice.ToTable(True, "DItemCode", "DQuantity", "DLineTotal", "DOutlet", "VatGourpSa")

                oDT_INTItemDetails_Final = GroupBy("DItemCode", "DQuantity", "DLineTotal", "DOutlet", "VatGourpSa", oDT_INTItemDetails)

                For Each dr As DataRow In oDT_INTItemDetails_Final.Rows
                    Dim result() As DataRow = oDT_Salesorder.Select("ItemCode = '" & dr("DItemCode").ToString.Trim & "'")
                    If result.Count > 0 Then
                        For Each row As DataRow In result
                            iLineNo = row("LineNum")
                            oSalesOrder.Lines.SetCurrentLine(iLineNo)
                        Next
                    Else
                        iLineNo = -1
                        If oSalesOrder.Lines.Count > 0 Then
                            oSalesOrder.Lines.Add()
                        End If
                    End If

                    oSalesOrder.Lines.ItemCode = dr("DItemCode").ToString.Trim
                    oSalesOrder.Lines.Quantity = dr("Quantity").ToString.Trim
                    oSalesOrder.Lines.LineTotal = dr("LineTotal").ToString.Trim
                    oSalesOrder.Lines.VatGroup = dr("VatGourpSa").ToString.Trim
                    oSalesOrder.Lines.WarehouseCode = dr("DOutlet").ToString.Trim
                Next

                lRetCode = oSalesOrder.Update()
                If lRetCode <> 0 Then
                    sErrDesc = oDICompany.GetLastErrorDescription
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    Salesorder = RTN_ERROR
                Else
                    Console.WriteLine("Completed with SUCCESS", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                    Salesorder = RTN_SUCCESS
                End If
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Salesorder = RTN_ERROR
        Finally
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oSalesOrder)
        End Try
    End Function

End Module
