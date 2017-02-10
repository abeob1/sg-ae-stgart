Module modARCreditMemo

    Function AR_CreditMemo(ByRef oDICompany As SAPbobsCOM.Company, _
                              ByVal sDocEntry As String, ByVal dIncomeDate As Date, _
                               ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim lRetCode As Long
        Dim oARCreditmemo As SAPbobsCOM.Documents
        Dim oARInvoice As SAPbobsCOM.Documents
        Dim sPayDocEntry As String = String.Empty

        oARCreditmemo = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
        oARInvoice = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)


        Try
            sFuncName = "AR_CreditMemo"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If oARInvoice.GetByKey(sDocEntry) Then

                oARCreditmemo.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInvoices
                oARCreditmemo.CardCode = oARInvoice.CardCode
                oARCreditmemo.NumAtCard = oARInvoice.NumAtCard
                oARCreditmemo.DocDate = oARInvoice.DocDate
                oARCreditmemo.TaxDate = oARInvoice.TaxDate
                oARCreditmemo.DocDueDate = oARInvoice.DocDueDate
                oARCreditmemo.DocType = oARInvoice.DocType

                For imjs As Integer = 0 To oARInvoice.Lines.Count - 1
                    oARInvoice.Lines.SetCurrentLine(imjs)
                    oARCreditmemo.Lines.BaseEntry = sDocEntry
                    oARCreditmemo.Lines.BaseLine = oARInvoice.Lines.LineNum
                    oARCreditmemo.Lines.BaseType = 13
                    '---------------------- Batch Information
                    For ibount As Integer = 0 To oARInvoice.Lines.BatchNumbers.Count - 1
                        oARCreditmemo.Lines.BatchNumbers.SetCurrentLine(ibount)
                        oARCreditmemo.Lines.BatchNumbers.BatchNumber = oARInvoice.Lines.BatchNumbers.BatchNumber
                        oARCreditmemo.Lines.BatchNumbers.Quantity = oARInvoice.Lines.BatchNumbers.Quantity
                        oARCreditmemo.Lines.BatchNumbers.Add()
                    Next
                    oARCreditmemo.Lines.Add()
                Next
                lRetCode = oARCreditmemo.Add()
                If lRetCode <> 0 Then
                    sErrDesc = oDICompany.GetLastErrorDescription
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
                    AR_CreditMemo = RTN_ERROR
                Else
                    Console.WriteLine("Completed with SUCCESS", sFuncName)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                    AR_CreditMemo = RTN_SUCCESS
                End If
            End If
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            AR_CreditMemo = RTN_ERROR
        Finally
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARCreditmemo)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oARInvoice)
        End Try
    End Function

End Module
