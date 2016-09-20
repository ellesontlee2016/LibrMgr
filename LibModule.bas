Attribute VB_Name = "LibModule"
Public Const ID_SECTIONMGR = 1000
Public Const ID_TEMPLATEMGR = 1001
Public Const ADDDELAY = 0
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
        (ByVal hwndCaller As Long, ByVal pszFile As String, _
        ByVal uCommand As Long, ByVal dwData As Long) As Long

Public Const HH_DISPLAY_TOPIC = &H0
Public Const HH_SET_WIN_TYPE = &H4
Public Const HH_GET_WIN_TYPE = &H5
Public Const HH_GET_WIN_HANDLE = &H6

' Display string resource ID or text in a popupwin.
Public Const HH_DISPLAY_TEXT_POPUP = &HE
' Display mapped numeric value in dwdata
Public Const HH_HELP_CONTEXT = &HF
' Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU
Public Const HH_TP_HELP_CONTEXTMENU = &H10
' Text pop-up help, similar to WinHelp's HELP_WM_HELP
Public Const HH_TP_HELP_WM_HELP = &H11

Public myAppForTemp As Object  'Word.Application
Public Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
 
Public Declare Function GetLongPathName Lib "kernel32" Alias _
    "GetLongPathNameA" (ByVal lpszShortPath As String, _
    ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
    
    
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, _
ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long



Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal ApphProcess As Long, _
ByVal uExitCode As Long) As Long

Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public myName As String
Const MAX_PATH = 260
Public defaultDocType As String
Const PASSWORD = "djei@%22cw3" '"!Gh@xyz#PROPGEN@CORS!!"

'Public myRegConAgent As New AuditUtil.RegistryAndConnection
'Public myChecker As New AuditUtil.ExpriationCheckAgent
'Public mySysInfo As New AuditUtil.UserKeyMain



Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Declare Function GetTempFileName Lib "kernel32" _
        Alias "GetTempFileNameA" (ByVal lpszPath As String, _
        ByVal lpPrefixString As String, ByVal wUnique As _
        Long, ByVal lpTempFileName As String) As Long
        
        
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal _
    lpFileName As String, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, ByVal NoSecurity As Long, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
    

Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, _
    lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, _
    lpLastWriteTime As FILETIME) As Long
    
    
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As _
    FILETIME, lpSystemTime As SYSTEMTIME) As Long
    
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As _
    FILETIME, lpLocalFileTime As FILETIME) As Long

Public Const GENERIC_READ = &H80000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const INVALID_HANDLE_VALUE = -1
Public bconvertEn As Boolean

Public bTempActive As Boolean
Public bSectionActive As Boolean
Public bTempUnload As Boolean
Public bSectionUnload As Boolean

Public bconvertAdo As Boolean
Public bco
  Public numberList As New Collection
  
 'Types Definition
Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'API Declarations
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long


'Constants Definition
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_EVENT = &H1
Private Const KEY_NOTIFY = &H10
Private Const READ_CONTROL = &H20000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const REG_BINARY = 3
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_EXPAND_SZ = 2
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
Private Const REG_NOTIFY_CHANGE_LAST_SET = &H4
Private Const REG_NOTIFY_CHANGE_NAME = &H1
Private Const REG_NOTIFY_CHANGE_SECURITY = &H8
Private Const REG_OPTION_BACKUP_RESTORE = 4
Private Const REG_OPTION_CREATE_LINK = 2
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_OPTION_RESERVED = 0
Private Const REG_OPTION_VOLATILE = 1
Private Const REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
Private Const REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)
'Public Enum STD_HANDLES
'  STD_INPUT = -10&
'  STD_OUTPUT = -11&
'  STD_ERROR = -12&
'End Enum

 
 
'Console input

Public Declare Function GetStdHandle _
  Lib "kernel32" _
(ByVal nStdHandle As Long) As Long


Public Declare Function WriteConsole _
  Lib "kernel32" Alias "WriteConsoleA" _
(ByVal hConsoleOutput As Long _
, ByVal lpBuffer As Any _
, ByVal nNumberOfCharsToWrite As Long _
, lpNumberOfCharsWritten As Long _
, lpReserved As Any _
) As Long

Public Declare Function ReadConsole _
  Lib "kernel32" Alias "ReadConsoleA" _
(ByVal hConsoleInput As Long _
, ByVal lpBuffer As String _
, ByVal nNumberOfCharsToRead As Long _
, lpNumberOfCharsRead As Long _
, lpReserved As Any _
) As Long






Public Function CRead() As String
  Dim sUserInput As String * 256
  Dim lUserInput As Long
  Dim lReturn As Long
  Dim sRead As String
  Dim hConsoleIn As Long
  hConsoleIn = GetStdHandle(STD_INPUT)

  lReturn = ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
  
  If lReturn > 0 Then
    lUserInput = InStr(sUserInput, Chr$(0))
    If lUserInput > 2 Then
      sRead = Left$(sUserInput, lUserInput - 3)
    Else
      sRead = sUserInput
    End If
  End If
  CRead = sRead

End Function
 'Delete e Registry Key with all his contained Values

Public Function FindLongTempPath() As String
    Dim sRet As String
    sRet = String$(MAX_PATH, 0)
    lRet = GetLongPathName(FindTemp(), sRet, Len(sRet))
    FindLongTempPath = Left(sRet, lRet)
End Function




Public Function getTempRoot() As String
  Dim ret As String
  ret = UCase(FindLongTempPath())

  getTempRoot = Left(ret, InStr(1, ret, "LOCAL SETTINGS") - 1)
End Function


 Public Function FindTemp() As String
    Dim result&, buff$
    buff = Space$(MAX_PATH)
    result = GetTempPath(Len(buff), buff)
    FindTemp = Left$(buff, result)
End Function
  
'  Public Sub updateRecord(Od As Double, otype As String, keepStyle As String, Desc As String, doc() As Byte, secName As String)

  
 ' Dim i As Integer
 ' Dim seclist As New Collection
' If DataEnvironment2.cnnSection.State <> adStateOpen Then
 ' getpwdConn DataEnvironment2.cnnSection
 ' DataEnvironment2.cnnSection.Open
'  End If
 ' DataEnvironment2.cmdUpdateRecord Od, otype, keepStyle, Desc, doc, secName, myName, Now
 
 ' If DataEnvironment2.cnnSection.State = adStateOpen Then
 ' DataEnvironment2.cnnSection.Close
 ' End If
 
'End Sub


Public Function OpenDoc(ByRef app As Object, filename As String) As Object
    Set retDoc = app.Documents.Open(filename) 'Add on Dec 6 2006
    Dim Spath As String
    Spath = app.Options.DefaultFilePath(8) 'wdStartupPath)
    'app.Run "StartUpMacro"
    
    retDoc.SaveAs filename
    
    
    retDoc.Application.Visible = True
    retDoc.Application.Activate
    retDoc.Application.ShowMe
    Set myAppForTemp = app
    Set OpenDoc = retDoc
End Function

Public Function openWordDoc(ByRef app As Object, filename As String) As Object ' Word.Application, FileName As String) As Document
 Dim retDoc As Object ' Document
 'Dim app As Word.Application
' Dim rets As Object
 'On Error GoTo HandleNull
 'rets = app.VisualBasic
 'Set openWordDoc = OpenDoc(app, filename)
   
 'Exit Function
   
'HandleNull:

 
On Error Resume Next
    If app Is Nothing Then
        Set app = GetObject(, "Word.Application")
        Debug.Print Err.Number
        If Err.Number <> 0 Then
            Set app = CreateObject("Word.Application")
            Err.Clear
        End If
        End If
        
Set openWordDoc = OpenDoc(app, filename)
    'End If
    'myWord1.Visible = True
   
  ''  If InStr(FileName, ".dot") <> 0 Then

  ''      Set retDoc = app.Documents.Open(FileName) '.Add(FileName, True, wdTypeTemplate, True)
  '' Else
   ''     Set retDoc = app.Documents.Add(FileName)
  ''  End If
 
    'app.ActiveDocument.name = filename
    
End Function
Public Function choseDBForExpCheck() As String
  Dim retstr As String
  Dim tobereplace, replacewith As String
  tobereplace = "PQuote\Data\PQDB.mdb"
  replacewith = "User\PQUsage.mdb"
  
  retstr = choseDB(False)
 ' retstr = replace(retstr, "PQDB.mdb", "PQX.mdb")
  retstr = replace(retstr, tobereplace, replacewith)
  choseDBForExpCheck = retstr
End Function


Public Function choseDB(Optional bIsCompany As Boolean = True) As String
    Dim constr As String
    Dim tmpstr As String
    Dim dbFullpath As String
    
    Dim value As String
    If bIsCompany Then
        value = GetRegValue(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\CorsPro\Settings", "Directory_Company")
    Else
        value = GetRegValue(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\CorsPro\Settings", "Directory_Root")

    End If
    If Right(value, 1) <> "\" Then
    value = value & "\"
    End If
    tmpstr = GetRegValue(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\CorsPro\Settings", "SubDir_PropGenDB")
    If tmpstr <> "" Then

    If Left(tmpstr, 0) = "\" Then
    tmpstr = Right(tmpstr, Len(tmpstr) - 1)
    End If
    dbFullpath = value & tmpstr
    Else
    dbFullpath = value & "PropGen\Data\PropDB.mdb"
    End If
    If bIsCompany Then
      DataEnvironment2.cnnSection.ConnectionString = replaceDataBase(DataEnvironment2.cnnSection.ConnectionString, "Data Source=" & dbFullpath)
    End If
      
    choseDB = dbFullpath
End Function


Public Function getApp() As Object ' Application
    Set getApp = myAppForTemp
End Function

Public Function getFileDateTime(filename As String) As Date
    Dim FileSys As Object
    Dim myFile As Object
    Set FileSys = CreateObject("System.FileSystemObject")
    Set myFile = FileSys.Open(filename)
    getFileDateTime = myFile.DateLastModified()
    
   End Function
   
Public Sub deleteFile(filename As String)
  On Error Resume Next
    Dim FileSys As Object
    Dim myFile As Object
    Set FileSys = CreateObject("System.FileSystemObject")
    FileSys.deleteFile filename
    
End Sub
   
Public Sub synOrderInfo()
 Dim index As Integer
 Dim rec As Recordset
 For index = 1 To numberList.Count
    numberList.Remove 1
 Next index
      If DataEnvironment2.cnnSection.State <> adStateOpen Then
      getpwdConn DataEnvironment2.cnnSection
      DataEnvironment2.cnnSection.Open
      End If
      
       DataEnvironment2.cmdGetOrderInfo
       Set rec = DataEnvironment2.rscmdGetOrderInfo
       
      While Not rec.EOF
        numberList.Add CStr(rec!Order_Number)
        rec.MoveNext
      Wend
      
      
     
      If DataEnvironment2.cnnSection.State = adStateOpen Then
      DataEnvironment2.cnnSection.Close
      End If
End Sub


Public Function isDuplicateOrderNo(no As Double) As Boolean
    Dim ret As Boolean
    Dim i As Integer
    ret = False
     
    synOrderInfo
    
    For i = 1 To numberList.Count
     If CDbl(numberList.item(i)) = no Then
       isDuplicateOrderNo = True
     End If
    Next i
    isDuplicateOrderNo = False
End Function

Public Function openToDB(Path As String) As Byte()
   Dim filename As String
   Dim fh As Integer
   Dim flen As Long
   Dim buffer() As Byte
   filename = Path
   If filename <> "" Then
   fh = FreeFile
    Open filename For Binary As #fh
    flen = LOF(fh)
    size = flen
    ReDim buffer(flen - 1) As Byte
    Get fh, , buffer
    Close #fh
    openToDB = buffer
    End If
   
End Function

Public Function OpenToDBNewTemplate(fullname As String) As Byte()

    Dim app As Object
    Dim namelist() As String
    Dim Doc As Object 'Word.Document
    Dim nameout As String
    'Dim name As String
    'Dim filename As String
    'namelist = Split(fullname, "\")
    'filename = namelist(UBound(namelist))
    'path = Left(fullname, Len(fullname) - Len(filename) - 1)
    
    nameout = CreateTempFile()
    
    'name = path & "\" & filename
    'nameout = path & "\_" & filename
    
    On Error Resume Next
    Set app = GetObject(, "Word.Application")
    Debug.Print Err.Number
    If Err.Number <> 0 Then
        Set app = CreateObject("Word.Application")
        Err.Clear
    End If
    Set Doc = app.Documents.Add(fullname)
    app.ActiveWindow.Visible = False


    'If InStr(app.name, "2007") <> 0 Then
    If isWord2007OrAbove(app.Path) Then
      app.ActiveDocument.SaveFormt = "*.dot"
      bIsWord2007 = True
      If InStr(nameout, ".docx") <> 0 Then
      addLog "Rename the *.dotx to *.dot"
       nameout = replace(nameout, ".dotx", ".dot")
      ElseIf InStr(nameout, ".dot") = 0 Then
      nameout = nameout & ".dot"
      End If
    End If
    addLog "Alway save it as wordtemplate format"

    app.ActiveDocument.SaveAs nameout, wdFormatTemplate_cp
    app.ActiveDocument.Close True
    
    OpenToDBNewTemplate = openToDB(nameout)
    deleteFile nameout
    
End Function

Public Function isWord2007OrAbove(Path As String) As Boolean
    Dim mPath As String
    Dim version As Integer
    mPath = UCase(Path)
    Dim index As Integer
    index = InStr(mPath, "\OFFICE")
    mPath = Right(mPath, Len(mPath) - index - 6)
    If mPath <> "" Then
    version = CInt(mPath)
    isWord2007OrAbove = version >= 12
    Else
    isWord2007OrAbove = False
    End If
    
End Function


Public Function OpenToDBNew(fullname As String) As Byte()
    Dim app As Object
    Dim namelist() As String
    Dim Doc As Object 'Word.Document
    Dim nameout As String
    'Dim name As String
    'Dim filename As String
    'namelist = Split(fullname, "\")
    'filename = namelist(UBound(namelist))
    'path = Left(fullname, Len(fullname) - Len(filename) - 1)
    
    nameout = CreateTempFile()
    
    'name = path & "\" & filename
    'nameout = path & "\_" & filename
    
    On Error Resume Next
    Set app = GetObject(, "Word.Application")
    Debug.Print Err.Number
    If Err.Number <> 0 Then
        Set app = CreateObject("Word.Application")
        Err.Clear
    End If

    Set Doc = app.Documents.Add(fullname)
    app.ActiveWindow.Visible = False
    'If InStr(app.name, "2007") <> 0 Then
    If isWord2007OrAbove(app.Path) Then
      app.ActiveDocument.SaveFormt = "*.doc"
      bIsWord2007 = True
      If InStr(nameout, ".docx") <> 0 Then
       nameout = replace(nameout, ".docx", ".doc")
       addLog "Rename the *.docx to *.doc"
      ElseIf InStr(nameout, ".doc") = 0 Then
      nameout = nameout & ".doc"
      End If
    End If
  
    addLog "Alway save it as word document format"
    app.ActiveDocument.SaveAs nameout, wdFormatDocument_cp
    If ADDDELAY Then
      Sleep 4000
    End If
    
    app.ActiveDocument.Close True
    OpenToDBNew = openToDB(nameout)
    deleteFile nameout
    
End Function

Public Function CreateTempFile() As String
    Dim TempDir$, result&, buff$
    Dim tmp As String
    TempDir = FindTemp
    If TempDir = "" Then Exit Function
    buff = Space$(MAX_PATH)
    result = GetTempFileName(TempDir, "~VB", 0&, buff)
    If result = 0 Then Exit Function

    result = InStr(1, buff, Chr(0))
    If result > 0 Then
      tmp = Left$(buff, result - 1)
    Else
      tmp = buff
    End If
    
    CreateTempFile = tmp 'left$(tmp, Len(tmp) - 4) & ".doc"
End Function





' Retrieve the Create date, Modify (write) date and Last Access date of
' the specified file. Returns True if successful, False otherwise.

Function GetFileTimeInfo(ByVal filename As String, Optional CreateDate As Date, _
    Optional ModifyDate As Date, Optional LastAccessDate As Date) As Boolean

    Dim hFile As Long
    Dim ftCreate As FILETIME
    Dim ftModify As FILETIME
    Dim ftLastAccess As FILETIME
    Dim ft As FILETIME
    Dim st As SYSTEMTIME
    
    ' open the file, exit if error
    hFile = CreateFile(filename, GENERIC_READ, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If hFile = INVALID_HANDLE_VALUE Then Exit Function
    
    ' read date information
    If GetFileTime(hFile, ftCreate, ftLastAccess, ftModify) Then
        ' non zero means successful
        GetFileTimeInfo = True
        
        ' convert result to date values
        ' first, convert UTC file time to local file time
        FileTimeToLocalFileTime ftCreate, ft
        ' then convert to system time
        FileTimeToSystemTime ft, st
        ' finally, make up the Date value
        CreateDate = DateSerial(st.wYear, st.wMonth, _
            st.wDay) + TimeSerial(st.wHour, st.wMinute, _
            st.wSecond) + (st.wMilliseconds / 86400000)
        
        ' do the same for the ModifyDate
        FileTimeToLocalFileTime ftModify, ft
        FileTimeToSystemTime ft, st
        ModifyDate = DateSerial(st.wYear, st.wMonth, _
            st.wDay) + TimeSerial(st.wHour, st.wMinute, _
            st.wSecond) + (st.wMilliseconds / 86400000)
        ' and for LastAccessDate
        FileTimeToLocalFileTime ftLastAccess, ft
        FileTimeToSystemTime ft, st
        LastAccessDate = DateSerial(st.wYear, st.wMonth, _
            st.wDay) + TimeSerial(st.wHour, st.wMinute, _
            st.wSecond) + (st.wMilliseconds / 86400000)
    End If
    
    ' close the file, in all cases
    CloseHandle hFile

End Function

Public Sub BubbleSortNumbers(iArray As Variant)

    Dim lLoop1 As Long
    Dim lLoop2 As Long
    Dim lTemp As Long
    
    For lLoop1 = UBound(iArray) To LBound(iArray) Step -1
      
        For lLoop2 = LBound(iArray) + 1 To lLoop1
          If iArray(lLoop2 - 1) > iArray(lLoop2) Then
              lTemp = iArray(lLoop2 - 1)
              iArray(lLoop2 - 1) = iArray(lLoop2)
              iArray(lLoop2) = lTemp
             
             '-----------------------------
             'Required for the speed Test;
             'comment out for real use
             'update the iterations label
              Bcnt = Bcnt + 1
              DoEvents
              If SkipFlag Then Exit Sub
             '----------------------------
              
          End If
        
        Next lLoop2
    
    Next lLoop1

End Sub

Public Function getXMLOfficeFiles() As Boolean
 On Error GoTo Error
      If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
      End If
      DataEnvironment2.cmdGetXMLFiles
       getXMLOfficeFiles = (DataEnvironment2.rscmdGetXMLFiles!UseXMLOfficeFiles = "Y")
       
      'New add code on Feb 14 2005, has not been tested yet
     If DataEnvironment2.cnnSection.State = adStateOpen Then
       DataEnvironment2.cnnSection.Close
      End If
      Exit Function
Error:
      If DataEnvironment2.cnnSection.State = adStateOpen Then
       DataEnvironment2.cnnSection.Close
      End If
     
getXMLOfficeFiles = False
 End Function

Public Function replaceDataBase(ByVal org As String, replace As String) As String
 Dim index1, index2 As Integer
 Dim ret As String
   index1 = InStr(1, org, "Data Source", vbTextCompare)
   
   If index1 <> 0 Then
    index2 = InStr(index1, org, ";", vbTextCompare)
    ret = Left(org, index1 - 1)
    If index2 <> -1 Then
        
        ret = ret & replace & Right(org, Len(org) - index2 + 1)
    Else
    ret = ret & replace
    
    End If
   Else
   ret = org & ";" & replace
   End If
   
   replaceDataBase = ret
End Function


Function StripItem(startStrg As String, parser As String) As String
'this takes a string separated by the chr passed in Parser,
'splits off 1 item, and shortens startStrg so that the next
'item is ready for removal.

   Dim c As Integer
   Dim item As String
   
   c = 1
   
   Do
   
      If Mid(startStrg, c, 1) = parser Then
      
         item = Mid(startStrg, 1, c - 1)
         startStrg = Mid(startStrg, c + 1, Len(startStrg))
         StripItem = item
         Exit Function
      End If
      
      c = c + 1
   
   Loop

End Function

Public Function getSecuredconnStr(connstr As String, pwd As String) As String
 If InStr(1, connstr, "Password") = 0 Then
    getSecuredconnStr = connstr & ";Jet OLEDB:Database Password=" & pwd
 Else
    getSecuredconnStr = connstr
 End If

End Function

Public Function getpwdconnStr(connstr As String) As String
 getpwdconnStr = getSecuredconnStr(connstr, PASSWORD)

End Function
Public Function getSecuredConn(cnn As Connection, pwd As String) As Connection
    cnn.Properties("Jet OLEDB:Database Password") = pwd
    
    Set getSecuredConn = cnn
End Function

    Public Sub getpwdConn(cnn As Connection)
Set cnn = getSecuredConn(cnn, PASSWORD)
End Sub


Public Function getClientName() As String

If myName = "" Then

    If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
    End If
    DataEnvironment2.cmdGetClientName
    myName = DataEnvironment2.rscmdGetClientName!RecName
    If DataEnvironment2.cnnSection.State = adStateOpen Then
    
        DataEnvironment2.cnnSection.Close
    End If
   
 End If
 
getClientName = myName
End Function

Public Function isSameRecSource(secName As String) As Boolean
     isSameRecSource = (myName = getRecSourceName(secName))
End Function
Public Function getRecSourceName(secName As String) As String
On Error GoTo Error
      If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
      End If
      DataEnvironment2.cmdGetRecSourceInfo secName
       getRecSourceName = DataEnvironment2.rscmdGetRecSourceInfo!RecSource
       
      'New add code on Feb 14 2005, has not been tested yet
     If DataEnvironment2.cnnSection.State = adStateOpen Then
        
       DataEnvironment2.cnnSection.Close
      End If
     Exit Function
Error:
      If DataEnvironment2.cnnSection.State = adStateOpen Then
       DataEnvironment2.cnnSection.Close
      End If
 End Function
 
 Public Function getDefaultDocType() As String
 On Error GoTo Error
      If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
      End If
      DataEnvironment2.cmdGetDefaultDocType
       getDefaultDocType = DataEnvironment2.rscmdGetDefaultDocType!defaultDocType
       
      'New add code on Feb 14 2005, has not been tested yet
     If DataEnvironment2.cnnSection.State = adStateOpen Then
       DataEnvironment2.cnnSection.Close
      End If
      Exit Function
Error:
      If DataEnvironment2.cnnSection.State = adStateOpen Then
       DataEnvironment2.cnnSection.Close
      End If
     
getDefaultDocType = "SOW"
 End Function


Public Function populateDocType() As Collection
Dim mycol As New Collection
If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
      End If
      
      On Error GoTo Error
    
    
        DataEnvironment2.cmdGetDocTypeList
        
    While Not DataEnvironment2.rscmdGetDocTypeList.EOF
        
        def = DataEnvironment2.rscmdGetDocTypeList!doctype
        mycol.Add (def)
        DataEnvironment2.rscmdGetDocTypeList.MoveNext
   
    Wend
      If DataEnvironment2.cnnSection.State = adStateOpen Then
        DataEnvironment2.cnnSection.Close
        
      End If
     Set populateDocType = mycol
      
Error:
      If DataEnvironment2.cnnSection.State = adStateOpen Then
      DataEnvironment2.cnnSection.Close
      End If
     Set populateDocType = mycol
End Function


Public Function getEDUtilPath() As String
   Dim edpath As String
   edpath = GetRegValue(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\CorsPro\Settings", "Directory_Root") & "\PQuote\EDUtil.exe"
   getEDUtilPath = """" & edpath & """"
End Function

Public Function isExpiredNew() As Boolean
On Error GoTo ErrorHandle
    Dim lngReturnCode As Long
    Dim strShellCommand As String
    Dim rname As String
    Dim pm As String
    Dim result As String
    Dim evalue As String
    Randomize
    evalue = "E" & CStr(Int((Rnd * 999) + 1))
    rname = "N" & CStr(Int((Rnd * 999) + 1))
    pm = "Software\PCDot\A" & CStr(Int((Rnd * 9) + 1))
    strShellCommand = getEDUtilPath() & " -op ep -vn " & rname & " -pr cu -pm " & pm & " -ev " & evalue
    Set oShell = CreateObject("WSCript.shell")
    oShell.run "cmd /C " & strShellCommand, 0, True
    result = GetRegValue(HKEY_CURRENT_USER, pm, rname)
    isExpiredNew = (result <> evalue)
    
    DeleteRegistryKey HKEY_CURRENT_USER, pm
    Exit Function
ErrorHandle:
    isExpiredNew = True
End Function

