Attribute VB_Name = "LogUtil"
Public SHOW_DEBUG As Integer ' 0 disable the debug, 1: enable debug to Applicationlog.txt
Public Const wdFormatDocument_cp = 0
Public Const wdFormatDocument97_cp = 0
Public Const wdFormatDocumentDefault_cp = 16
Public Const wdFormatDOSText_cp = 4
Public Const wdFormatDOSTextLineBreaks_cp = 5
Public Const wdFormatEncodedText_cp = 7
Public Const wdFormatFilteredHTML_cp = 10
Public Const wdFormatFlatXML_cp = 19
Public Const wdFormatFlatXMLMacroEnabled_cp = 20
Public Const wdFormatFlatXMLTemplate_cp = 21
Public Const wdFormatFlatXMLTemplateMacroEnabled_cp = 22
Public Const wdFormatHTML_cp = 8
Public Const wdFormatPDF_cp = 17
Public Const wdFormatRTF_cp = 6
Public Const wdFormatTemplate_cp = 1
Public Const wdFormatTemplate97_cp = 1
Public Const wdFormatText_cp = 2
Public Const wdFormatTextLineBreaks_cp = 3
Public Const wdFormatUnicodeText_cp = 7
Public Const wdFormatWebArchive_cp = 9
Public Const wdFormatXML_cp = 11
Public Const wdFormatXMLDocument_cp = 12
Public Const wdFormatXMLDocumentMacroEnabled_cp = 13
Public Const wdFormatXMLTemplate_cp = 14
Public Const wdFormatXMLTemplateMacroEnabled_cp = 15
Public Const wdFormatXPS_cp = 18
    
Public Type LogRecord
  sDate As String * 8
  sTime As String * 8
  sErrorNumber As String * 6
  sConditionCode As String * 1
  sModuleName As String * 20
  sLogMessage As String * 83
 
  sCrLf As String * 2
End Type

Private Type HeaderRecord
  nTotalRecords As Integer
  nNextRecordPosition As Integer
  nFullFlag As Integer
  sFiller As String * 122
End Type

Public Sub addLog(msg As String)
'TxtList.Text = TxtList.Text & msg & vbCrLf
WriteLogRecord msg, 0, ""

End Sub

Public Sub WriteLogRecord(sLogMessage As String, _
                          nErrorNumber As Integer, _
                          sConditionCode As String)
' writes a message to the Log file
If SHOW_DEBUG <> 1 Then
 Exit Sub
End If

SimpleLogWrite sLogMessage
Exit Sub


Dim nLogFileNum As Integer
Dim CurrentHeader As HeaderRecord
Dim CurrentLogRecord As LogRecord
Dim nCurrentRecordNum As Integer
Dim nMaxRecords As Integer
Dim sTestLog As String
Dim sModule As String
Dim sErrorNum As String
Dim sLogFileName As String

'fix up log information for module name, error
sModule = app.EXEName
sErrorNum = Trim$(Str$(nErrorNumber))
sLogFileName = app.path & "\ApplicationLog.log"

' See if log file exists, create if it does not
' (Note: Production code should have more robust
'  technique to check for file existence.)
sTestLog = Dir$(sLogFileName)        ' returns null string if file not found
If sTestLog = "" Then                ' if no log file present
    CreateLogFile sLogFileName, 1000 ' make a new one, 1000 records
End If

' open the log file
nLogFileNum = FreeFile
Open sLogFileName For Random As nLogFileNum Len = 128

'    get header record and extract needed variables
Get nLogFileNum, 1, CurrentHeader
nMaxRecords = CurrentHeader.nTotalRecords
nCurrentRecordNum = CurrentHeader.nNextRecordPosition
If nCurrentRecordNum > nMaxRecords Then
        nCurrentRecordNum = 2
        CurrentHeader.nFullFlag = 1
End If

'    set information for writing log record

CurrentLogRecord.sDate = Date$
CurrentLogRecord.sTime = Time$
CurrentLogRecord.sModuleName = sModule
CurrentLogRecord.sConditionCode = sConditionCode
CurrentLogRecord.sLogMessage = sLogMessage
CurrentLogRecord.sErrorNumber = sErrorNum
CurrentLogRecord.sCrLf = vbCrLf

'    write the log record
Put nLogFileNum, nCurrentRecordNum, CurrentLogRecord

'    update header
CurrentHeader.nNextRecordPosition = nCurrentRecordNum + 1
RSet CurrentHeader.sFiller = vbCrLf
Put nLogFileNum, 1, CurrentHeader

'    we're done - close log and quit
Close nLogFileNum

End Sub

Public Sub CreateLogFile(sLogFileName As String, nNumRecordsToUse As Integer)

Dim CurrentHeader As HeaderRecord
Dim nLogFileNum As Integer

CurrentHeader.nTotalRecords = nNumRecordsToUse
CurrentHeader.nNextRecordPosition = 2
RSet CurrentHeader.sFiller = vbCrLf

nLogFileNum = FreeFile
Open sLogFileName For Random As nLogFileNum Len = 128
Put nLogFileNum, 1, CurrentHeader
Close nLogFileNum

End Sub



Function GetLogMessage(nStartItem As Integer, sLoggedDate As String, sLoggedTime As String, sLoggedModule As String, sLoggedConditionCode As String, sLoggedMessage As String, sErrorNum As String) As Integer
Static CurrentRecNum As Integer
Dim nLogFileNum As Integer, CurrentHeader As HeaderRecord
Dim CurrentLogRecord As LogRecord, nRecordToRead As Integer

' Retrieves log messages. Can be called repeatedly. If nStartItem is specified
' then a new set of requests is assumed. Is nStartItem is 0 then the next item is retrieved.
' If nStartItem is positive, the number of messages to begin is counted from
' the oldest message in the log. If nStartItem is negative, the starting message
' is counted backwards from the newest message in the log.

' The function returns 0 as long as a valid log message was read. When the end of the log
' is reached (the newest message has been retrieved) or when an invalid nStartItem was
' passed, the function returns -1

' Function assumes log files exists. Production code should have a check
' that this is true.

If nStartItem = 0 And CurrentRecNum = 0 Then
    ' no current start number has been specified, so exit
    GetLogMessage = -1
    Exit Function
End If

GetLogMessage = 0              ' initialize function to 0


' open log file and retrieve number of records and current oldest record.
nLogFileNum = FreeFile
Dim sLogFileName As String
sLogFileName = app.path & "\TestLog.log"

Open sLogFileName For Random As nLogFileNum Len = 128
Get nLogFileNum, 1, CurrentHeader

If Abs(nStartItem) > CurrentHeader.nTotalRecords Then      ' no such record
    GetLogMessage = -1
    CurrentRecNum = 0
    Exit Function
End If

' set the current logical record to read based on nStartItem
Select Case nStartItem
    Case Is > 0
        CurrentRecNum = nStartItem
    Case Is < 0
        CurrentRecNum = CurrentHeader.nTotalRecords + nStartItem + 1
    Case 0
        CurrentRecNum = CurrentRecNum + 1
        If CurrentRecNum > CurrentHeader.nTotalRecords Then
            GetLogMessage = -1
            CurrentRecNum = 0
            Exit Function
        End If
End Select

' set the current physical record to read based on CurrentRecNum (the current logical record)
' need to check to see if the log file is full to determine how to make the calculation

If CurrentHeader.nFullFlag = 1 Then
    If CurrentRecNum <= CurrentHeader.nTotalRecords - CurrentHeader.nNextRecordPosition + 1 Then
        nRecordToRead = CurrentHeader.nNextRecordPosition + CurrentRecNum - 1
    Else
        nRecordToRead = CurrentRecNum - CurrentHeader.nTotalRecords + CurrentHeader.nNextRecordPosition
    End If
Else
    If CurrentRecNum > CurrentHeader.nNextRecordPosition - 1 Then
        GetLogMessage = -1
        CurrentRecNum = 0
        Exit Function
    End If
    nRecordToRead = CurrentRecNum + 1
End If

' read the log file record and place data in return arguments
Get nLogFileNum, nRecordToRead, CurrentLogRecord
Close nLogFileNum
sLoggedDate = CurrentLogRecord.sDate
sLoggedTime = CurrentLogRecord.sTime
sLoggedModule = CurrentLogRecord.sModuleName
sLoggedConditionCode = CurrentLogRecord.sConditionCode
sLoggedMessage = CurrentLogRecord.sLogMessage
sErrorNum = CurrentLogRecord.sErrorNumber

End Function


Public Sub SimpleLogWrite(ByVal sLogMessage As String)
  Dim nFileNum As Integer
  On Error GoTo handle:
  nFileNum = FreeFile
  Open app.path & "\ApplicationLog.log" For Append As nFileNum
  sLogMessage = Date$ & " - " & Time$ & " - " & sLogMessage
  Print #nFileNum, sLogMessage
handle:
  
  Close nFileNum
End Sub


