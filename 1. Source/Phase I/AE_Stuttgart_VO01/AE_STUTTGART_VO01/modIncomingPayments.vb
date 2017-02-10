Module modIncomingPayments

    Function AR_IncomingPayment(ByRef oDVPayment As DataView, ByVal oDICompany As SAPbobsCOM.Company, _
                                ByVal sDocEntry As String, ByVal dIncomeDate As Date, _
                                ByVal sPOSNumber As String, ByVal sWhsCode As String, _
                               ByVal sCardCode As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim lRetCode As Long
        Dim oIncomingPayment As SAPbobsCOM.Payments
        Dim sPayDocEntry As String = String.Empty

        Try
            sFuncName = "AR_IncomingPayment"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oIncomingPayment = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

            Dim sCreditCard As String = String.Empty

            oIncomingPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
            oIncomingPayment.CardCode = CStr(sCardCode)
            oIncomingPayment.DocDate = dIncomeDate
            oIncomingPayment.DueDate = dIncomeDate
            oIncomingPayment.TaxDate = dIncomeDate
            oIncomingPayment.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments

            If sDocEntry <> "" Then
                oIncomingPayment.Invoices.DocEntry = sDocEntry
                oIncomingPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                oIncomingPayment.Invoices.Add()
            End If

            For Each drv In oDVPayment
                If drv("PaymentAmount").ToString.Trim = 0.0 Then Continue For

                oIncomingPayment.CreditCards.CreditCard = drv("CreditCard").ToString.Trim
                oIncomingPayment.CreditCards.CreditType = SAPbobsCOM.BoRcptCredTypes.cr_Regular
                oIncomingPayment.CreditCards.CardValidUntil = "01/12/9999"
                oIncomingPayment.CreditCards.CreditCardNumber = "1234" 'drv("CreditNumber").ToString.Trim
                oIncomingPayment.CreditCards.VoucherNum = sWhsCode & "-" & CDate(dIncomeDate).ToString("yyMMdd") & "-" & sPOSNumber
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Amount : " & CDbl(drv("PaymentAmount").ToString.Trim), sFuncName)
                oIncomingPayment.CreditCards.CreditSum = CDbl(drv("PaymentAmount").ToString.Trim)
                oIncomingPayment.CreditCards.Add()
            Next

            oIncomingPayment.CashSum = 0

            Console.WriteLine("Attempting to Add ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add  ", sFuncName)
            lRetCode = oIncomingPayment.Add()

            If lRetCode <> 0 Then
                sErrDesc = oDICompany.GetLastErrorDescription
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)

                AR_IncomingPayment = RTN_ERROR
            Else

                Console.WriteLine("Completed with SUCCESS " & sDocEntry, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                AR_IncomingPayment = RTN_SUCCESS

            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            AR_IncomingPayment = RTN_ERROR

        Finally
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment)

            oIncomingPayment = Nothing
        End Try
    End Function

    Function Downpayment_IncomingPayment(ByRef oDVPayment As DataView, ByVal oDICompany As SAPbobsCOM.Company, _
                              ByVal sDocEntry As String, ByVal dIncomeDate As Date, _
                              ByVal sPOSNumber As String, ByVal sWhsCode As String, _
                             ByVal sCardCode As String, ByRef sErrDesc As String) As Long

        Dim sFuncName As String = String.Empty
        Dim lRetCode As Long
        Dim oIncomingPayment As SAPbobsCOM.Payments
        Dim sPayDocEntry As String = String.Empty

        Try
            sFuncName = "AR_IncomingPayment"
            Console.WriteLine("Starting Function", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oIncomingPayment = oDICompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)

            Dim sCreditCard As String = String.Empty

            oIncomingPayment.DocType = SAPbobsCOM.BoRcptTypes.rCustomer
            oIncomingPayment.CardCode = CStr(sCardCode)
            oIncomingPayment.DocDate = dIncomeDate
            oIncomingPayment.DueDate = dIncomeDate
            oIncomingPayment.TaxDate = dIncomeDate
            oIncomingPayment.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments

            If sDocEntry <> "" Then
                oIncomingPayment.Invoices.DocEntry = sDocEntry
                oIncomingPayment.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_DownPayment
                oIncomingPayment.Invoices.Add()
            End If

            For Each drv In oDVPayment
                If drv("PaymentAmount").ToString.Trim = 0.0 Then Continue For

                oIncomingPayment.CreditCards.CreditCard = drv("CreditCard").ToString.Trim
                oIncomingPayment.CreditCards.CreditType = SAPbobsCOM.BoRcptCredTypes.cr_Regular
                oIncomingPayment.CreditCards.CardValidUntil = "01/12/9999"
                oIncomingPayment.CreditCards.CreditCardNumber = "1234" 'drv("CreditNumber").ToString.Trim
                oIncomingPayment.CreditCards.VoucherNum = sWhsCode & "-" & CDate(dIncomeDate).ToString("yyMMdd") & "-" & sPOSNumber
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Amount : " & CDbl(drv("PaymentAmount").ToString.Trim), sFuncName)
                oIncomingPayment.CreditCards.CreditSum = CDbl(drv("PaymentAmount").ToString.Trim)
                oIncomingPayment.CreditCards.Add()
            Next

            oIncomingPayment.CashSum = 0

            Console.WriteLine("Attempting to Add ", sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add  ", sFuncName)
            lRetCode = oIncomingPayment.Add()

            If lRetCode <> 0 Then
                sErrDesc = oDICompany.GetLastErrorDescription
                Call WriteToLogFile(sErrDesc, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)

                Downpayment_IncomingPayment = RTN_ERROR
            Else

                Console.WriteLine("Completed with SUCCESS " & sDocEntry, sFuncName)
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS.", sFuncName)
                Downpayment_IncomingPayment = RTN_SUCCESS

            End If

        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            Downpayment_IncomingPayment = RTN_ERROR

        Finally
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Releasing the Objects", sFuncName)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oIncomingPayment)

            oIncomingPayment = Nothing
        End Try
    End Function


End Module
