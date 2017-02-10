Option Explicit On

Imports System.IO
Imports System.Configuration



Module modLog

    Public Structure CompanyDefault

        Public p_sServerName As String
        Public p_sLicServerName As String
        Public p_sDataBaseName As String
        Public p_sIntDBName As String
        Public p_sDBUserName As String
        Public p_sDBPassword As String
        Public p_sSAPUserName As String
        Public p_sSAPPassword As String
        Public p_sCardCode As String
        Public p_sSQLType As Int16
        Public p_sLogDir As String
        Public p_sDebug As String
        Public p_sRoundingItem As String
        Public p_sZeroTax As String
        Public p_iIntegDays As Integer

    End Structure


    Public Const RTN_SUCCESS As Int16 = 1
    Public Const RTN_ERROR As Int16 = 0
    ' Debug Value Variable Control
    Public Const DEBUG_ON As Int16 = 1
    Public Const DEBUG_OFF As Int16 = 0

    ' Global variables group
    Public p_iDebugMode As Int16
    Public p_iErrDispMethod As Int16
    Public p_iDeleteDebugLog As Int16
    Public p_oCompDef As CompanyDefault

    ' Public p_oCompany As SAPbobsCOM.Company
    Public p_oDataView As DataView = Nothing
    Public P_sConString As String = String.Empty
    Public p_sSAPConnString As String = String.Empty
    Public p_oCompany As SAPbobsCOM.Company
    Public P_sCardCode As String = String.Empty

    '***************************************
    'Name       :   modLog
    'Descrption :   Contains function for log errors and Application related information
    'Author     :   JOHN
    'Created    :   MAY 2014
    '***************************************

    Private Const MAXFILESIZE_IN_MB As Int16 = 5 '(2 MB)
    Private Const LOG_FILE_ERROR As String = "ErrorLog"
    Private Const LOG_FILE_ERROR_ARCH As String = "ErrorLog_"
    Private Const LOG_FILE_DEBUG As String = "DebugLog"
    Private Const LOG_FILE_DEBUG_ARCH As String = "DebugLog_"
    Private Const FILE_SIZE_CHECK_ENABLE As Int16 = 1
    Private Const FILE_SIZE_CHECK_DISABLE As Int16 = 0

    Public Function WriteToLogFile(ByVal strErrText As String, ByVal strSourceName As String, Optional ByVal intCheckFileForDelete As Int16 = 1) As Long

        ' **********************************************************************************
        '   Function   :    WriteToLogFile()
        '   Purpose    :    This function checks if given input file name exists or not
        '
        '   Parameters :    ByVal strErrText As String
        '                       strErrText = Text to be written to the log
        '                   ByVal intLogType As Integer
        '                       intLogType = Log type (1 - Log ; 2 - Error ; 0 - None)
        '                   ByVal strSourceName As String
        '                       strSourceName = Function name calling this function
        '                   Optional ByVal intCheckFileForDelete As Integer
        '                       intCheckFileForDelete = Flag to indicate if file size need to be checked before logging (0 - No check ; 1 - Check)
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    JOHN
        '   Date       :    MAY 2014
        ' **********************************************************************************

        Dim oStreamWriter As StreamWriter = Nothing
        Dim strFileName As String = String.Empty
        Dim strArchFileName As String = String.Empty
        Dim strTempString As String = String.Empty
        Dim lngFileSizeInMB As Double

        Try
            strTempString = Space(IIf(Len(strSourceName) > 30, 0, 30 - Len(strSourceName)))
            strSourceName = strTempString & strSourceName
            strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText


            strFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_ERROR & ".log"
            strArchFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_ERROR_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"


            'strFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_ERROR & ".log"
            'strArchFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_ERROR_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"

            If intCheckFileForDelete = FILE_SIZE_CHECK_ENABLE Then
                If File.Exists(strFileName) Then
                    lngFileSizeInMB = (FileLen(strFileName) / 1024) / 1024
                    If lngFileSizeInMB >= MAXFILESIZE_IN_MB Then
                        'If intCheckDeleteDebugLog=1 then remove all debug_log file
                        If p_iDeleteDebugLog = 1 Then
                            For Each sFileName As String In Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), LOG_FILE_DEBUG_ARCH & "*")
                                File.Delete(sFileName)
                            Next
                        End If
                        File.Move(strFileName, strArchFileName)
                    End If
                End If
            End If
            oStreamWriter = File.AppendText(strFileName)
            oStreamWriter.WriteLine(strErrText)
            WriteToLogFile = RTN_SUCCESS
        Catch exc As Exception
            WriteToLogFile = RTN_ERROR
        Finally
            If Not IsNothing(oStreamWriter) Then
                oStreamWriter.Flush()
                oStreamWriter.Close()
                oStreamWriter = Nothing
            End If
        End Try

    End Function

    Public Function WriteToLogFile_Debug(ByVal strErrText As String, ByVal strSourceName As String, Optional ByVal intCheckFileForDelete As Int16 = 1) As Long
        ' **********************************************************************************
        '   Function   :    WriteToLogFile_Debug()
        '   Purpose    :    This function checks if given input file name exists or not
        '
        '   Parameters :    ByVal strErrText As String
        '                       strErrText = Text to be written to the log
        '                   ByVal intLogType As Integer
        '                       intLogType = Log type (1 - Log ; 2 - Error ; 0 - None)
        '                   ByVal strSourceName As String
        '                       strSourceName = Function name calling this function
        '                   Optional ByVal intCheckFileForDelete As Integer
        '                       intCheckFileForDelete = Flag to indicate if file size need to be checked before logging (0 - No check ; 1 - Check)
        '   Return     :    0 - FAILURE
        '                   1 - SUCCESS
        '   Author     :    JOHN
        '   Date       :    MAY 2013
        '   Changes    : 
        '                   
        ' **********************************************************************************

        Dim oStreamWriter As StreamWriter = Nothing
        Dim strFileName As String = String.Empty
        Dim strArchFileName As String = String.Empty
        Dim strTempString As String = String.Empty
        Dim lngFileSizeInMB As Double
        Dim iFileCount As Integer = 0

        Try
            strTempString = Space(IIf(Len(strSourceName) > 30, 0, 30 - Len(strSourceName)))
            strSourceName = strTempString & strSourceName
            strErrText = "[" & Format(Now, "MM/dd/yyyy HH:mm:ss") & "]" & "[" & strSourceName & "] " & strErrText

            If String.IsNullOrEmpty(p_oCompDef.p_sLogDir) Then
                p_oCompDef.p_sLogDir = System.IO.Directory.GetCurrentDirectory()
            End If

            strFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_DEBUG & ".log"
            strArchFileName = p_oCompDef.p_sLogDir & "\" & LOG_FILE_DEBUG_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"


            'strFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_DEBUG & ".log"
            'strArchFileName = System.IO.Directory.GetCurrentDirectory() & "\" & LOG_FILE_DEBUG_ARCH & Format(Now(), "yyMMddHHMMss") & ".log"

            If intCheckFileForDelete = FILE_SIZE_CHECK_ENABLE Then
                If File.Exists(strFileName) Then
                    lngFileSizeInMB = (FileLen(strFileName) / 1024) / 1024
                    If lngFileSizeInMB >= MAXFILESIZE_IN_MB Then
                        'If intCheckDeleteDebugLog=1 then remove all debug_log file
                        If p_iDeleteDebugLog = 1 Then
                            For Each sFileName As String In Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), LOG_FILE_DEBUG_ARCH & "*")
                                File.Delete(sFileName)
                            Next
                        End If
                        File.Move(strFileName, strArchFileName)
                    End If
                End If
            End If
            oStreamWriter = File.AppendText(strFileName)
            oStreamWriter.WriteLine(strErrText)
            WriteToLogFile_Debug = RTN_SUCCESS
        Catch exc As Exception
            WriteToLogFile_Debug = RTN_ERROR
        Finally
            If Not IsNothing(oStreamWriter) Then
                oStreamWriter.Flush()
                oStreamWriter.Close()
                oStreamWriter = Nothing
            End If
        End Try

    End Function

  

End Module
