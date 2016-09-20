VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStoreDocNew 
   Caption         =   "Section Manager"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   Icon            =   "frmStoreDocNew.frx":0000
   LinkTopic       =   "frmStoreDocNew"
   MDIChild        =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   7800
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Edit"
      Height          =   255
      HelpContextID   =   1000
      Left            =   5160
      TabIndex        =   4
      Top             =   150
      Width           =   852
   End
   Begin VB.CommandButton cmdDelete1 
      Caption         =   "Delete"
      Height          =   255
      HelpContextID   =   1000
      Left            =   6840
      TabIndex        =   3
      Top             =   150
      Width           =   852
   End
   Begin VB.CommandButton cmdBatchAdd 
      Caption         =   "Add"
      Height          =   255
      HelpContextID   =   1000
      Left            =   6000
      TabIndex        =   2
      Top             =   150
      Width           =   852
   End
   Begin VB.CommandButton btnUp 
      Caption         =   "Move Up"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   1095
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move Down"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   150
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      HelpContextID   =   1000
      Left            =   120
      Negotiate       =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   315
      Left            =   3600
      Top             =   2640
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   3
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=;Persist Security Info=False;"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=;Persist Security Info=False;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "Admin"
      Password        =   ""
      RecordSource    =   $"frmStoreDocNew.frx":1CCA
      Caption         =   "SectionTable"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lbmm 
      Caption         =   "Label7"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmStoreDocNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim size As Long
Dim bClose As Boolean
Dim bUpdate As Boolean
Dim defaultDocPath As String
Dim diff_H, diff_W As Integer
Dim o_H, o_W As Integer
Dim diff_BL As Integer
Dim old_H As Integer
Dim currentSelectIndex As Integer
Dim frmCon As frmConfirm
Dim myParent As frmLibManager
Dim myfrm As frmUpdate
Dim myName As String
Dim SectionFileLen As Integer
Dim currentLocalIndex As Integer
Dim currentLocalBeginIndex As Integer
Dim LocalRowSize As Integer
Dim currentSelName As String
Dim TrackIndex As Integer
Dim firstSelectIndex As Integer
Dim currentBookmark As Integer
Dim currentBeginIndex As Integer

    Dim myCurrentName, prevName, nextName As String

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)
Const REG_SZ As Long = 1
Const REG_DWORD As Long = 4

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const MY_FUN = 800
Const ERROR_NONE = 0
Const ERROR_BADDB = 1
Const ERROR_BADKEY = 2
Const ERROR_CANTOPEN = 3
Const ERROR_CANTREAD = 4
Const ERROR_CANTWRITE = 5
Const ERROR_OUTOFMEMORY = 6
Const ERROR_INVALID_PARAMETER = 7
Const ERROR_ACCESS_DENIED = 8
Const ERROR_INVALID_PARAMETERS = 87
Const ERROR_NO_MORE_ITEMS = 259

Const KEY_ALL_ACCESS = &H3F

Const REG_OPTION_NON_VOLATILE = 0

Const KEY_WRITE = &H20006

Const REG_BINARY = 3

Const KEY_READ = &H20019



Const REG_OPENED_EXISTING_KEY = &H2

Const REG_EXPAND_SZ = 2
Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234

Dim bStart As Boolean
Dim bUpdateDocOK As Boolean




Const DDL_READWRITE = &H0
Const LB_DIR = &H18D
Dim sPattern As String
Dim counterx As Integer
Dim myWord1 As Object 'Word.Application
Dim myDoc As Object 'Word.Document
Dim iDaysUntilDelete As Integer
Dim offset As Integer

Dim bIsWord2007 As Boolean
Dim bForcedByMDIP As Boolean

' Create a registry key, then close it
' Returns True if the key already existed, False if it was created


Public Function GetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
  Dim handle As Long
  Dim resLong As Long
  Dim resString As String
  Dim resBinary() As Byte
  Dim length As Long
  Dim retVal As Long
  Dim valueType As Long
  GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
  If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
      Exit Function
  End If
  length = 1024
  ReDim resBinary(0 To length - 1) As Byte
  retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), length)
  If retVal = ERROR_MORE_DATA Then
      ReDim resBinary(0 To length - 1) As Byte
      retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), length)
  End If
  Select Case valueType
      Case REG_DWORD
          CopyMemory resLong, resBinary(0), 4
          GetRegistryValue = resLong
      Case REG_SZ, REG_EXPAND_SZ
          resString = Space$(length - 1)
          CopyMemory ByVal resString, resBinary(0), length - 1
          GetRegistryValue = resString
      Case REG_BINARY
          If length <> UBound(resBinary) + 1 Then
              ReDim Preserve resBinary(0 To length - 1) As Byte
          End If
          GetRegistryValue = resBinary()
      Case REG_MULTI_SZ
          resString = Space$(length - 2)
          CopyMemory ByVal resString, resBinary(0), length - 2
          GetRegistryValue = resString
      Case Else
          RegCloseKey handle
  End Select
  RegCloseKey handle
End Function

Public Function SetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, ByVal ValueName As String, value As Variant) As Boolean
  Dim handle As Long
  Dim lngValue As Long
  Dim strValue As String
  Dim binValue() As Byte
  Dim length As Long
  Dim retVal As Long
  If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then
      SetRegistryValue = False
      Exit Function
  End If
  Select Case VarType(value)
      Case vbInteger, vbLong
          lngValue = value
          retVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
      Case vbString
          strValue = value
          retVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, Len(strValue))
      Case vbArray + vbByte
          binValue = value
          length = UBound(binValue) - LBound(binValue) + 1
          retVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, binValue(LBound(binValue)), length)
      Case Else
          SetRegistryValue = False
          RegCloseKey handle
          Exit Function
  End Select
  If CBool(retVal) Then
   
   SetRegistryValue = False
   RegCloseKey handle
   Exit Function
  End If
  RegCloseKey handle
  SetRegistryValue = (retVal = 0)
End Function

Public Function CreateRegistryKey(lPredefinedKey As Long, sNewKeyName As String) As Long
  
   Dim hNewKey As Long         'handle to the new key
  Dim lRetVal As Long         'result of the RegCreateKeyEx function
  
   lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
      vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
      0&, hNewKey, lRetVal)
  CreateRegistryKey = lRetVal
  RegCloseKey (hNewKey)
End Function


Public Function getFileWithDefaultPath(filename As String) As String
On Error Resume Next
    Set myWord1 = GetObject(, "Word.Application")
    Debug.Print Err.Number
    If Err.Number <> 0 Then Set myWord1 = CreateObject("Word.Application")
    Err.Clear
    myWord1.Visible = True
    defaultDocPath = myWord1.Options.DefaultFilePath(8) 'wdDocumentsPath)
    If InStrRev(filename, "\") <> 0 Then
        filename = defaultDocPath & "\" & Right(filename, Len(filename) - InStrRev(filename, "\"))
    Else
        filename = defaultDocPath & "\" & filename
End If
    getFileWithDefaultPath = filename
End Function




''Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'currentSelectIndex = Me.Adodc1.Recordset.AbsolutePosition
'If Me.DataGrid1.Row >= 0 Then
'currentSelectIndex = Me.Adodc1.Recordset.AbsolutePosition

'End If
'Dim index As Integer
'index = Me.Adodc1.Recordset.AbsolutePosition

'Me.btnUp.Enabled = index <> 1
'Me.cmdDown.Enabled = index <> Me.Adodc1.Recordset.RecordCount
'Me.DataGrid1.SelStart = Me.Adodc1.Recordset.AbsolutePosition

'Me.btnUp.Enabled = offset <> 0
'Me.cmdDown.Enabled = offset <> Me.Adodc1.Recordset.RecordCount - 1

''End Sub



Private Sub btnSource_Click()
Dim buff() As Byte
 CommonDialog1.ShowOpen
' Me.txtSource.Text = CommonDialog1.filename
' Dim name As String
' name = Me.txtSource.Text
 
' name = Right(name, Len(name) - InStrRev(name, "\"))
' Me.txtName.Text = name
 synch

End Sub

Private Sub writeTOFile(filename As String, data() As Byte)

   Dim fh As Integer
   Dim flen As Long
   fh = FreeFile
   Open filename For Binary As #fh
   Put fh, , data
   Close #fh

End Sub
Private Sub AdjustView()
Dim oindex, viewsize As Integer
Dim cbIndex As Integer

  On Error GoTo handle
    oindex = Me.Adodc1.Recordset.Bookmark
    cbIndex = Me.DataGrid1.FirstRow
    viewsize = Me.DataGrid1.visibleRows
    If currentLocalBeginIndex >= cbIndex And currentLocalBeginIndex < cbIndex + viewsize Then
    Else
    Me.Adodc1.Recordset.Requery
         Me.Adodc1.Recordset.Move oindex - 1, 1

    DataGrid1_ColResize 0, 0
    
    End If
    
  
handle:
End Sub
Private Sub handleUpNew()
    Dim originalIndex As Integer
    Dim currentRowIndex As Integer
    Dim tmp As Object
    'currentRowIndex = Me.DataGrid1.Row
    Dim visibleRows As Integer
    currentRowIndex = Me.DataGrid1.Row
    Dim beginIndex As Integer
    DataGrid1.Visible = False
    AdjustView
    DataGrid1.Visible = True
  
    beginIndex = Me.DataGrid1.FirstRow
    currentLocalBeginIndex = beginIndex
    visibleRows = Me.DataGrid1.visibleRows
    originalIndex = Me.Adodc1.Recordset.Bookmark
    
    'Check if the selected row is in the view
    'and not in the first row, then no paging is change
    DataGrid1.Visible = False
     MoveRow originalIndex, originalIndex - 1
     Me.Adodc1.Recordset.Requery
     DataGrid1_ColResize 0, 0
     DataGrid1.Visible = True
     
     If visibleRows = originalIndex Then
      Me.Adodc1.Recordset.Move 0, 1
        Me.DataGrid1.Row = originalIndex - 2
        Me.btnUp.Enabled = originalIndex - 2 <> 0
        Me.cmdDown.Enabled = True
        Exit Sub
     End If
     
     If originalIndex >= beginIndex + visibleRows - 1 Then
             DataGrid1.Visible = False


          Me.Adodc1.Recordset.Move originalIndex - visibleRows, 1
          beginIndex = Me.DataGrid1.FirstRow
          Me.DataGrid1.Scroll 0, originalIndex - visibleRows
         Me.DataGrid1.Row = originalIndex - originalIndex + visibleRows - 2
           Me.btnUp.Enabled = originalIndex - 2 <> 0
        Me.cmdDown.Enabled = True
            DataGrid1.Visible = True


        Exit Sub
 
     End If
     
    If beginIndex - 1 > visibleRows Or beginIndex = 1 Then
     '   If currentRowIndex >= 1 Then
      DataGrid1.Visible = False
            Me.Adodc1.Recordset.Move beginIndex - 1, 1
            
                beginIndex = Me.Adodc1.Recordset.AbsolutePosition

            If currentRowIndex >= 1 Then
                Me.DataGrid1.Row = currentRowIndex - 1
            Else
                If beginIndex > visibleRows + 1 Then
                    Me.Adodc1.Recordset.Move beginIndex - visibleRows - 1, 1
                    Me.DataGrid1.Row = visibleRows - 1
                    If Me.DataGrid1.Row = 0 Then
                        'This is a trick way to solve this problem
                        'Move to the top and scroll up
                        
                        Me.Adodc1.Recordset.Move beginIndex - 2, 1
                        Me.DataGrid1.Scroll 0, -visibleRows + 2
                    Else
                    End If
                'Me.Adodc1.Recordset.Move beginIndex + visibleRows + 1, 1
                'Me.Adodc1.Recordset.Move beginIndex - visibleRows - 2, 1
                'Me.DataGrid1.Row = visibleRows - 1
                
                  '  Me.Adodc1.Recordset.AbsolutePosition = adPosBOF
                  Else
               ' Me.Adodc1.Recordset.Move 1, 1
               '  Me.DataGrid1.Row = beginIndex - 3

                End If
                End If
            DataGrid1.Visible = True
         Else
         DataGrid1.Visible = False

         Me.Adodc1.Recordset.Move beginIndex - 1, 1
         Me.DataGrid1.Row = currentRowIndex - 1
         DataGrid1.Visible = True
    
       End If
     Me.btnUp.Enabled = originalIndex - 2 <> 0
     Me.cmdDown.Enabled = True
    
End Sub


Private Sub handleDownNew()
    Dim originalIndex As Integer
    Dim currentRowIndex As Integer
    Dim tmp As Object
    'currentRowIndex = Me.DataGrid1.Row
    Dim visibleRows As Integer
    AdjustView
    currentRowIndex = Me.DataGrid1.Row
    Dim beginIndex As Integer
    beginIndex = Me.DataGrid1.FirstRow
    currentLocalBeginIndex = beginIndex
    visibleRows = Me.DataGrid1.visibleRows
    originalIndex = Me.Adodc1.Recordset.Bookmark
    'Check if the selected row is in the view
    'and not in the first row, then no paging is change
    If originalIndex + 1 <= Me.Adodc1.Recordset.RecordCount Then
    DataGrid1.Visible = False
        MoveRow originalIndex, originalIndex + 1
        Me.Adodc1.Recordset.Requery
        DataGrid1_ColResize 0, 0
         If currentRowIndex < visibleRows - 1 Then
            Me.Adodc1.Recordset.Move beginIndex - 1, 1
            Me.DataGrid1.Row = currentRowIndex + 1
            
            If beginIndex = 1 Or beginIndex - 1 > visibleRows Then
                Me.Adodc1.Recordset.Move beginIndex - 1, 1
                Me.DataGrid1.Row = currentRowIndex + 1
            Else
                Me.Adodc1.Recordset.Move beginIndex + currentRowIndex, 1
            End If
            Me.cmdDown.Enabled = beginIndex + currentRowIndex <> Me.Adodc1.Recordset.RecordCount - 1
        End If
        DataGrid1.Visible = True
     Else
        Me.cmdDown.Enabled = False
        
        Me.Adodc1.Recordset.MoveLast
   
    End If
    
    Me.btnUp.Enabled = True

    
    
End Sub

Private Sub handleUp()
    Dim originalIndex As Integer
    Dim currentRowIndex As Integer
    'currentRowIndex = Me.DataGrid1.Row
    Dim visibleRows As Integer
    visibleRows = Me.DataGrid1.visibleRows
   originalIndex = Me.Adodc1.Recordset.Bookmark
    If originalIndex = 1 Then
        cmdDown.Enabled = True
        btnUp.Enabled = True
        Exit Sub
    End If

    MoveRow originalIndex, originalIndex - 1
    Me.Adodc1.Recordset.Requery
    DataGrid1_ColResize 0, 0
    If originalIndex > visibleRows + currentLocalIndex Then
    Me.Adodc1.Recordset.Move originalIndex - currentLocalIndex - 2, 1 'originalIndex - 2, 1 ' (selindex - 2)
    DataGrid1.Row = currentLocalIndex
    TrackIndex = 1
    Else
       If TrackIndex < currentLocalIndex Then
       Me.Adodc1.Recordset.Move originalIndex - currentLocalIndex - 2 + TrackIndex, 1
       DataGrid1.Row = currentLocalIndex - TrackIndex
       TrackIndex = TrackIndex + 1
       Else
       Me.Adodc1.Recordset.Move originalIndex - 2, 1
      
       End If
    End If
    
 
   ' DataGrid1_Click

End Sub

Private Sub btnUp_Click()
'handleUp
Me.WindowState = 2
handleUpNew
Me.lbmm.Caption = CStr(Me.DataGrid1.FirstRow) & "," & CStr(Me.Adodc1.Recordset.Bookmark) & "," & CStr(Me.DataGrid1.visibleRows)
Exit Sub

''Dim selindex As Integer
''selindex = currentSelectIndex


    Dim originalIndex As Integer
    
    currentSelName = Me.Adodc1.Recordset!Section_Name
   originalIndex = Me.Adodc1.Recordset.Bookmark
    If originalIndex = 1 Then
        cmdDown.Enabled = True
        btnUp.Enabled = True
        Exit Sub
    End If

 '   originalIndex2 = Me.DataGrid1.Row
''Dim selindex As Integer
''selindex = currentSelectIndex
MoveRow originalIndex, originalIndex - 1


''MoveRow selindex, selindex - 1

Me.Adodc1.Recordset.Requery
Dim k As Integer
Dim tpmp As String
Dim sizerow As Integer

sizerow = findoutSize
'This portion is related to the refresh issue
'if the currentLocalIndex >=1 then, we swap and also record the currentLocalIndex

If currentLocalIndex >= 1 Then

    Me.Adodc1.Recordset.Move currentLocalBeginIndex - 1, 1 ' (selindex - 2)
    currentLocalBeginIndex = Me.DataGrid1.FirstRow
    Me.DataGrid1.Row = currentLocalIndex - 1

     ' If currentSelName <> Me.Adodc1.Recordset!Section_Name Then
     ' Me.DataGrid1.Row = currentLocalBeginIndex
     ' End If
    currentLocalIndex = Me.DataGrid1.Row
    Else
    If currentLocalBeginIndex >= 5 Then
    Me.Adodc1.Recordset.Move currentLocalBeginIndex - 5, 1
    currentLocalBeginIndex = Me.DataGrid1.FirstRow
    
     If currentLocalBeginIndex > sizerow + 1 Then
      Me.DataGrid1.Row = originalIndex - currentLocalBeginIndex + 3
     Else
     
     Me.Adodc1.Recordset.Move 0, 1
     currentLocalBeginIndex = Me.DataGrid1.FirstRow
     If originalIndex - currentLocalBeginIndex - 1 < 0 Then
     Me.DataGrid1.Row = 0
     Else
     
      Me.DataGrid1.Row = originalIndex - currentLocalBeginIndex - 1
      
      End If
      
      If Me.DataGrid1.Row < 0 Then
       Me.DataGrid1.Row = 0
       End If
      
      If Me.DataGrid1.Row <> originalIndex - currentLocalBeginIndex - 1 Then
      
      
      End If
      End If
      
     
     End If
     

    
    'If currentLocalBeginIndex >= 5 Then
   '  If currentLocalBeginIndex - 2 <= sizerow Then
   '      Me.Adodc1.Recordset.Move 0, 1
   '      Me.DataGrid1.Row = currentLocalBeginIndex - 1
     
   '  Else
     
  '    Me.Adodc1.Recordset.Move currentLocalBeginIndex - 5, 1
  '    currentLocalBeginIndex = Me.DataGrid1.FirstRow
  '    Me.DataGrid1.Row = originalIndex - currentLocalBeginIndex + 3
     
     'If originalIndex >= currentLocalBeginIndex Then
    
     ' Me.Adodc1.Recordset.Move currentLocalBeginIndex - 5, 1
    ' Me.DataGrid1.Row = originalIndex - currentLocalBeginIndex + 3
   '  If currentLocalBeginIndex - 1 <= sizerow Then
     'If currentSelName <> Me.Adodc1.Recordset!Section_Name Then
        ' Me.DataGrid1.Row = currentLocalBeginIndex - 2
    '     Me.Adodc1.Recordset.Move 0, 1
   '      Me.DataGrid1.Row = currentLocalBeginIndex - 1
    ' End If
    ' Else
     
       '  Me.Adodc1.Recordset.Move 0, 1
      '   Me.DataGrid1.Row = currentLocalBeginIndex - 1
     
     
     
   ' End If
    
      'Me.DataGrid1.Row = currentLocalIndex + 1
      currentLocalIndex = Me.DataGrid1.Row
 '   Else
    k = 0
  End If
  
   
 ' End If
  
    
DataGrid1_ColResize 0, 0
currentLocalBeginIndex = Me.DataGrid1.FirstRow

''Me.Adodc1.Recordset.Move selindex - 2, 1 ' (selindex - 2)
    DataGrid1.Visible = True
'DataGrid1.FirstRow = currentLocalBeginIndex
'DataGrid1.Row = currentLocalIndex - 1
'Me.DataGrid1.Row = currentLocalIndex - 1
'Me.DataGrid1.Row = selindex - 2
'DataGrid1_Click
'Me.DataGrid1.Row = Me.DataGrid1.FirstRow + Me.DataGrid1.Row - 1


End Sub

Private Function findoutSize() As Integer

Dim i As Integer
Dim current As Integer

current = Me.DataGrid1.Row
For i = 0 To 200
Me.DataGrid1.Row = i
If Me.DataGrid1.Row <> i Then
Me.DataGrid1.Row = current
findoutSize = i
Exit Function
End If
Next
Me.DataGrid1.Row = current
findoutSize = 1000
End Function

Private Sub cmdDown_Click()
Me.WindowState = 2
handleDownNew
Me.lbmm.Caption = CStr(Me.DataGrid1.FirstRow) & "," & CStr(Me.Adodc1.Recordset.Bookmark) & "," & CStr(Me.DataGrid1.visibleRows)

Exit Sub


    Dim originalIndex As Integer
    Dim originalIndex2 As Integer

    
    originalIndex = Me.Adodc1.Recordset.Bookmark
    originalIndex2 = Me.DataGrid1.Row

    
    If originalIndex >= Me.Adodc1.Recordset.RecordCount Then
        cmdDown.Enabled = True
        btnUp.Enabled = True
        Exit Sub
    End If

 '   originalIndex2 = Me.DataGrid1.Row
''Dim selindex As Integer
''selindex = currentSelectIndex

MoveRow originalIndex, originalIndex + 1

''MoveRow selindex, selindex + 1
'Me.Adodc1.Recordset.Save
'DataGrid1.Visible = False
'Me.Refresh
Me.Adodc1.Recordset.Requery
DataGrid1_ColResize 0, 0
'cmdSort_Click
''Me.Adodc1.Recordset.Move selindex, 1 ' (selindex - 2)
Me.Adodc1.Recordset.Move originalIndex, 1 ' (selindex - 2)

DataGrid1.Visible = True

'Me.DataGrid1.Row = selindex - 2
DataGrid1_Click
End Sub


Private Sub cmdAddDoc_Click()
    Dim buff() As Byte
    Dim docName As String
    
    Dim comm As Command
    On Error Resume Next
        CommonDialog1.CancelError = True
        CommonDialog1.ShowOpen
          
        If Err.Number <> 0 Then
        Exit Sub
        End If
     
    docName = CommonDialog1.filename
    docName = Right(docName, Len(docName) - InStrRev(docName, "\"))
     
    Dim name As String
    
    insertNewRecord Me.Adodc1.Recordset, docName
    On Error Resume Next
    Me.Adodc1.Refresh
    DataGrid1.ReBind
    
    Me.Adodc1.Recordset.MoveLast
    DataGrid1_ColResize 0, 0
      
End Sub




Private Sub insertNewRecord(rec As Recordset, name As String)
    Dim buff() As Byte

    'CalculateCounter True
    Dim mypars(6) As Parameter
    Dim mycomd As New Command
    buff = OpenToDBNew(name)
    bClose = False
    
    myName = getClientName()
    If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
    End If
    mycomd.ActiveConnection = DataEnvironment2.cnnSection
    
    mycomd.CommandText = "INSERT INTO Section_tbl(Section_Name, Order_Number, Object_Type,Word_Doc,Keep_Style,Description,RecSource) VALUES (?,?,?,?,?,?,?)"
   
    Set mypars(0) = mycomd.CreateParameter("Section_Name", adVariant, adParamInput, 40)
    Set mypars(1) = mycomd.CreateParameter("Order_Number", adVariant, adParamInput, 40)
    Set mypars(2) = mycomd.CreateParameter("Object_Type", adVariant, adParamInput, 40)
    Set mypars(3) = mycomd.CreateParameter("Word_Doc", adBSTR, adParamInput, size)
    Set mypars(4) = mycomd.CreateParameter("Keep_Style", adVariant, adParamInput, 40)
    Set mypars(5) = mycomd.CreateParameter("Description", adVariant, adParamInput, 255)
    Set mypars(6) = mycomd.CreateParameter("RecSource", adVariant, adParamInput, 50)
    mypars(0).value = rec!Section_Name
    mypars(1).value = rec!Order_Number
    mypars(2).value = rec!Object_Type
    mypars(3).value = buff
    mypars(4).value = rec!Keep_Style
    mypars(5).value = rec!Description
    mypars(6).value = myName
    
    mycomd.Parameters.Append mypars(0)
    mycomd.Parameters.Append mypars(1)
    mycomd.Parameters.Append mypars(2)
    mycomd.Parameters.Append mypars(3)
    mycomd.Parameters.Append mypars(4)
    mycomd.Parameters.Append mypars(5)
    mycomd.Parameters.Append mypars(6)
    
    mycomd.CommandType = adCmdText
    On Error GoTo myHandler
    
    mycomd.Execute
    bClose = True
    Exit Sub
    

myHandler:
MsgBox Err.Description
    MsgBox "Duplicate document name, please use different one", vbOKOnly, "Warning"
    If Not bClose Then
   End If

End Sub

Private Sub cmdAdd_Click()
    Dim index As Integer
    Dim dlgUpdate As New frmUpdate
    
    dlgUpdate.loadInfo Me.Adodc1.Recordset, CDbl(currentSelectIndex), True
    dlgUpdate.setButtonFocus True
    
    
    dlgUpdate.Show vbModal
    
    dlgUpdate.cmdUpdate.Enabled = False
    
    Me.Adodc1.Recordset.Requery
    Me.Adodc1.Refresh
    cmdSort_Click
    
    For index = 1 To currentSelectIndex - 1
        Me.Adodc1.Recordset.MoveNext
    Next
    
    If dlgUpdate.isAddedOK Then Me.Adodc1.Recordset.MoveNext
    
    If Me.Adodc1.Recordset.RecordCount > 0 Then
    cmdDelete1.Enabled = True
    End If
    
End Sub


Private Sub Insertdoc()
    
        frmUpdate.AddRec
   '     Else
   '     UpdateRec
   '     End If
        
   ' End If


End Sub
Public Sub AddRec(sec As String, orderNumber As Double, buff() As Byte, Optional doctype As String = "internal")
    Dim i As Integer
    Dim myName As String
    myName = getClientName()
      If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
      End If
      
      On Error GoTo Error
      addmsg "sec name=" & sec
      addmsg "sec name=" & sec
      addmsg "sec name=" & sec
      addmsg "sec name=" & sec
      addmsg "sec name=" & sec
      Dim infomsg As String
      addDebug "Sec=" & sec
      addDebug "ordernumber =" & orderNumber
      addDebug "defaultDocType=" & doctype
      addDebug "KeppAlive=Yes"
      addDebug "Recordsource=" & myName
      addDebug "Description=null"
    
        DataEnvironment2.cmdInsertDocNew sec, orderNumber, "Internal", doctype, buff, "Yes", "", myName, Now
    
      If DataEnvironment2.cnnSection.State = adStateOpen Then
        DataEnvironment2.cnnSection.Close
        
      End If
      
      Exit Sub
      
Error:
      If DataEnvironment2.cnnSection.State = adStateOpen Then
      DataEnvironment2.cnnSection.Close
      End If
      addDebug "Error: " & Err.Description
      MsgBox Err.Description, , "Warning"
End Sub


Public Sub AddRecExternal(sec As String, orderNumber As Double, keepStyle As String, doctype As String, Desc As String)
    Dim i As Integer
    myName = getClientName()
      If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
      End If
      
      On Error GoTo Error
        DataEnvironment2.cmdInsertDocNew sec, orderNumber, "External", doctype, Empty, "Yes", Desc, myName, CDate(Now)
      If DataEnvironment2.cnnSection.State = adStateOpen Then
      DataEnvironment2.cnnSection.Close
      End If
      Exit Sub
      
Error:
      If DataEnvironment2.cnnSection.State = adStateOpen Then
      DataEnvironment2.cnnSection.Close
      End If
      MsgBox Err.Description, , "Warning"
End Sub


Private Sub addmsg(msg As String)
  'Me.lbtemp.Caption = msg
End Sub

Private Sub cmdBatchAdd_Click()
  'working variables
   Dim c As Integer
   Dim y As Integer
   Dim sFile As String
   Dim startStrg As String
   Dim originalMark As Integer
   Dim tmp As String
    Dim i As Integer
    Dim buffer() As Byte
    Dim docName As String
    Dim secName As String
    Dim incD As Double
    Dim bInsertBefore As Boolean
    bInsertBefore = True
    On Error GoTo handle:
    
    Dim originalIndex As Integer
    Dim originalIndex2 As Integer
    originalIndex = Me.Adodc1.Recordset.Bookmark
    originalIndex2 = Me.DataGrid1.Row
    
addLog "Start Add Record Processing"
 
    
   ' If MsgBox("Do you want to insert the new record before this row?", vbYesNo, "Insertion Option Confirmation") = vbNo Then
   ' bInsertBefore = False
  '  End If
    
    originalMark = Me.DataGrid1.Row
addLog "Invoke getDefaultDocType"

    defaultDocType = getDefaultDocType()
    
    If frmCon Is Nothing Then
     Set frmCon = New frmConfirm
    End If
    
    frmCon.Show vbModal, myParent
    If Not frmCon.isOKClick() Then
       Exit Sub
    End If

    If Not frmCon.isInsertBefore() Then
    bInsertBefore = False
    End If
    
    ' Screen.MousePointer = vbHourglass

    If Not frmCon.isInternal() Then
        showProperties True, True
        Screen.MousePointer = vbHourglass
    

        If myfrm.isCanceled Then
           Screen.MousePointer = vbDefault

           Set myfrm = Nothing
           Exit Sub
        End If
        AddRecExternal myfrm.getSecName(), myfrm.getOrderNumber(), myfrm.getKeepStyle(), myfrm.getDocType(), myfrm.getDesc()
        Set myfrm = Nothing
        Me.cmdDelete1.Enabled = True
        Dim cursor As Integer
        
        Me.Adodc1.Recordset.Requery
        Me.Adodc1.Refresh
        cmdSort_Click
        cursor = currentSelectIndex
        For cursor = 1 To currentSelectIndex
            If cursor <= Me.Adodc1.Recordset.RecordCount Then
                Me.Adodc1.Recordset.MoveNext
            End If
        Next
       ' cmdSort_Click
    
    Else
    
 
  'dim an array to hold the files selected
   Dim FileArray() As String
    
  'set the max buffer large enough to retrieve multiple files
   CommonDialog1.MaxFileSize = 4096
   CommonDialog1.DefaultExt = "doc"
   CommonDialog1.DialogTitle = "Add Section(s)"
   If getXMLOfficeFiles() Then
                 CommonDialog1.Filter = "Word Document (*.doc;*.docx)|*.doc;*.docx"
                 CommonDialog1.filename = "*.doc;*.docx"
                 CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
                 CommonDialog1.ShowOpen     ' = 1
           
              
            If Right(CommonDialog1.filename, 5) = "*.doc" Or Right(CommonDialog1.filename, 6) = "*.docx" Then
        
                  Screen.MousePointer = vbDefault
                 Exit Sub
                 End If
                 
          
                If Right(CommonDialog1.filename, 4) <> ".doc" And Right(CommonDialog1.filename, 5) <> ".docx" Then
                   MsgBox "ERROR: Illegal file type; only DOC/DOCX files can be imported into the document library.", vbOKOnly, "Error"
                   Screen.MousePointer = vbDefault
                 Exit Sub
                 End If
                 
            Else
             CommonDialog1.Filter = "Word Document (*.doc)|*.doc"
             CommonDialog1.filename = "*.doc"
             CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
             CommonDialog1.ShowOpen     ' = 1
           
             If Right(CommonDialog1.filename, 5) = "*.doc" Then
        
                Screen.MousePointer = vbDefault
                Exit Sub

             End If
             
             
               If Right(CommonDialog1.filename, 4) <> ".doc" Then
                   MsgBox "ERROR: Illegal file type; only DOC files can be imported into the document library.", vbOKOnly, "Error"
                   Screen.MousePointer = vbDefault
                 Exit Sub
                 End If
                       
             
           End If
   startStrg = CommonDialog1.filename & Chr(0) & Chr(0)
   
   For c = 1 To Len(CommonDialog1.filename)
      
     'extract 1 item from the string
      sFile = StripItem(startStrg, Chr(0))
      
     'if nothing's there, we're done
      If sFile = "" Then Exit For
      
        'ReDim the filename array
        'to add the new file. FileArray(0) is either the
        'path (if more than 1 file selected), or the
        'fully qualified filename (if only 1 file selected).
         ReDim Preserve FileArray(0 To c - 1)
         FileArray(y) = LCase(sFile)
         
        'increment y by 1 for the next pass
         y = y + 1
      
      Next
      
    Dim DocLR30ErrorMsg   As String
    Dim DocsLR30ErrorMsg As String
    
    
    DocLR30ErrorMsg = "You attempted to import a document with a filename greater than " & CStr(SectionFileLen) & " characters." & vbCrLf & _
    "Please rename the document using " & CStr(SectionFileLen) & " characters or less, then re-import it."

    DocsLR30ErrorMsg = "You attempted to import one or more documents with filenames greater than " & CStr(SectionFileLen) & " characters." & vbCrLf & _
    "Please rename the documents so that their filenames are " & CStr(SectionFileLen) & " characters or less, then re-import them."
      
    If Not validateFileLen(FileArray) Then
      If UBound(FileArray) = 0 Then
      MsgBox DocLR30ErrorMsg, vbOKOnly, "Error"
      Exit Sub
      
      Else
            MsgBox DocsLR30ErrorMsg, vbOKOnly, "Error"
      Exit Sub

      End If
      
    End If
    
      
      
      
    Dim prefix As String
    c = UBound(FileArray)
    
    'Check the file name first
 addLog "Start check the selected document "
 
    Dim prvIndex As Double
    Dim nxtIndex As Double
    Dim currentIndex As Double
    
    myCurrentName = getSecName()
    prevName = getPreviousSecName()
    nextName = getNextSecName()
    
    currentIndex = getCurrentOrder(myCurrentName)
    Dim dist As Double

    
    If prevName <> "" Then
        prvIndex = getCurrentOrder(prevName)
    Else
        prvIndex = currentIndex - 0.1
    End If

    If nextName <> "" Then
        nxtIndex = getCurrentOrder(nextName)
    Else
        nxtIndex = currentIndex + 0.1
    End If


    If bInsertBefore Then
        dist = (currentIndex - prvIndex) / (c + 2)
    Else
        dist = (nxtIndex - currentIndex) / (c + 2)
    End If

 
    If c = 0 Then
    incD = 2
    
    
        docName = FileArray(0)
        If docName <> "" Then
         secName = Mid(docName, InStrRev(docName, "\") + 1, InStrRev(docName, ".") - InStrRev(docName, "\") - 1)
         
addLog "Start OpenToDBNew "
         buffer = OpenToDBNew(FileArray(0))
         
addLog "(1) add record with secName=" & secName & " with order=" & CDbl(originalIndex - incD * 0.001)
         
     
         If bInsertBefore Then
          AddRec secName, CDbl(currentIndex - dist), buffer, frmCon.getDocType()
          'AddRec secName, CDbl(originalIndex - incD * 0.001), buffer
         Else
          
          AddRec secName, CDbl(currentIndex + dist), buffer, frmCon.getDocType()
         ' AddRec secName, CDbl(originalIndex + incD * 0.001), buffer
         End If
         
         Me.cmdDelete1.Enabled = True
        End If
    Else
    
        
    
    prefix = FileArray(0)
    For i = 1 To c
    incD = CDbl(i)
        docName = FileArray(i)
        If docName <> "" Then
         secName = Mid(docName, InStrRev(docName, "\") + 1, InStrRev(docName, ".") - InStrRev(docName, "\") - 1)
         buffer = OpenToDBNew(prefix & "\" & docName)

         If bInsertBefore Then
            'AddRec secName, CDbl(originalIndex - incD * 0.001), buffer
            AddRec secName, CDbl(currentIndex - dist * i), buffer, frmCon.getDocType()
         Else
            AddRec secName, CDbl(currentIndex + dist * i), buffer, frmCon.getDocType()
         End If

         Me.cmdDelete1.Enabled = True
        End If
    Next i
    End If
   ' Me.Adodc1.Recordset.Requery
    'Me.Adodc1.Refresh
addLog "Invoke Sort "
  ' cmdSort_Click
  
    Me.Adodc1.Refresh
    DataGrid1_ColResize 0, 0

 '   Me.DataGrid1.Row = originalIndex2
 '   Me.Adodc1.Recordset.Move originalIndex - 1, 1

addLog "Adjuest DataGrid view "
   
    If bInsertBefore Then

        Me.DataGrid1.Row = originalIndex2 - 1
        Me.Adodc1.Recordset.Move originalIndex - 1, 1


   '  Me.DataGrid1.Row = currentSelectIndex - Me.DataGrid1.FirstRow
     
     Else
         ' Me.DataGrid1.Row = currentSelectIndex - Me.DataGrid1.FirstRow + 1
        Me.DataGrid1.Row = originalIndex2 + 1
        Me.Adodc1.Recordset.Move originalIndex + c - 1, 1

     End If
     
addLog "Finish adjusting DataGrid view"
     
  End If
      Screen.MousePointer = vbDefault

    Exit Sub
handle:
     Screen.MousePointer = vbDefault

End Sub

Private Function getCurrentOrder(ByVal sectionname As String) As Double
    Dim i As Integer
      If DataEnvironment2.cnnSection.State <> adStateOpen Then
      getpwdConn DataEnvironment2.cnnSection
      DataEnvironment2.cnnSection.Open
 
        DataEnvironment2.cmdGetOrder sectionname
      End If
      getCurrentOrder = DataEnvironment2.rscmdGetOrder(0)
     
      If DataEnvironment2.cnnSection.State = adStateOpen Then
      DataEnvironment2.cnnSection.Close
      End If
      
End Function
Private Function validateFileLen(file() As String) As Boolean
Dim secName As String
Dim bPass As Boolean
bPass = True
Dim i, c As Integer
c = UBound(file)
 If c = 0 Then
          secName = file(0)
          secName = Mid(secName, InStrRev(secName, "\") + 1, InStrRev(secName, ".") - InStrRev(secName, "\") - 1)
          bPass = Len(secName) <= 30
 Else
     
     For i = 1 To c - 1
          secName = file(i)
          secName = Mid(secName, InStrRev(secName, "\") + 1, InStrRev(secName, ".") - InStrRev(secName, "\") - 1)
          bPass = bPass And Len(secName) <= 30
     Next i
 End If
 validateFileLen = bPass
 

End Function

Private Function getOrderValue(index As Integer) As Double
On Error GoTo handle:
If index <= Me.Adodc1.Recordset.RecordCount - 1 Then
'Me.Adodc1.Recordset.MoveFirst

Me.Adodc1.Recordset.Move index, 1
getOrderValue = Me.Adodc1.Recordset!Order_Number
Exit Function
Else
getOrderValue = -1
End If
handle:
getOrderValue = -1
End Function

Private Function setOrderValue(value As Double, index As Integer)

Me.Adodc1.Recordset.Move index - 1, 1
Me.Adodc1.Recordset!Order_Number = value

End Function


Private Sub cmdDelete1_Click()
    Dim i As Integer
    myName = getClientName()
    Dim bDeleteError As Boolean
    Dim bAlreadyPopConfirmation As Boolean
    
    Dim Retmsg As String
    Dim selindex As Integer
    Dim AnyArray() As Long
    Dim offset As Integer
    Dim selectlist As New Collection
    Dim tmp As String
    Screen.MousePointer = vbHourglass
    If Me.Adodc1.Recordset.EOF Then
    Screen.MousePointer = vbDefault
     Exit Sub
     End If
    Dim totalDelete As Integer
    bAlreadyPopConfirmation = False
    totalDelete = 0
    For i = 0 To Me.DataGrid1.SelBookmarks.Count - 1
    
   'Me.Adodc1.Recordset.MoveFirst
        Me.Adodc1.Recordset.Move (Me.DataGrid1.SelBookmarks.item(i) - 1), 1
        
        If isSameRecSource(Me.Adodc1.Recordset!Section_Name) Then
            selectlist.Add Me.DataGrid1.SelBookmarks.item(i)
            totalDelete = totalDelete + 1
        End If
    Next i
    offset = 0
    
    If Me.DataGrid1.SelBookmarks.Count = 0 Then
       Screen.MousePointer = vbDefault
       MsgBox "To delete doc sections, you must select one or more ENTIRE rows", vbOKOnly, "Section Manager"
       Exit Sub
    End If
    ReDim AnyArray(totalDelete - 1)

    For i = 1 To selectlist.Count
        AnyArray(i - 1) = CLng(selectlist.item(i))
    Next i
    
    Retmsg = "Are you sure you want to delete "
   
   
    
    BubbleSortNumbers AnyArray
    
Dim bDeleteCurrent As Boolean
Dim currentDeleteIndex As Integer
bDeleteCurrent = True
offset = 0
   
    For i = 1 To selectlist.Count
    

    ' moveRecordSet CInt(selectlist.item(i)) - 1 - offset
    currentDeleteIndex = CInt(AnyArray(i - 1)) - 1 - offset
         moveRecordSet currentDeleteIndex
        'If MsgBox("Do you want to remove " & Me.Adodc1.Recordset!Section_Name, vbYesNo, "Confirmation") = vbYes Then
        If DataEnvironment2.cnnSection.State <> adStateOpen Then
            getpwdConn DataEnvironment2.cnnSection
            DataEnvironment2.cnnSection.Open
        End If
       Dim title As String
       
       Dim msg As String
       Dim secName As String
       secName = Me.Adodc1.Recordset!Section_Name
       secName = Trim(secName)
       msg = getconfirmationMsg("DocSectionsByPart", secName)
       If msg <> "" Then
            msg = msg & " and" & vbCrLf & getconfirmationMsg("DocSectionsByItem", secName)
       Else
            msg = getconfirmationMsg("DocSectionsByItem ", secName)
       End If
       
       If msg <> "" Then
       msg = "Do you want to delete this record with Section_Name='" & secName & "'?" & vbCrLf & "When you delete this record, you will aslo be deleting " & vbCrLf _
              & msg & "." & vbCrLf
              '& : Only the record with RecSource='" & myName & "' will be deleted!"
             Else
             
        msg = "Do you want to delete this record with Section_Name='" & secName & "'?" & vbCrLf
        '& "Note: Only the record with RecSource='" & myName & "' will be deleted!"
            
             End If
             

       
      ' If Not bAlreadyPopConfirmation Then
     '       If Me.DataGrid1.SelBookmarks.Count = 1 Then
     '          Retmsg = Retmsg & "this section?" 'Me.Adodc1.Recordset!Section_Name & " ?"
      '          title = "Delete Section"
     '       Else
    '            If Me.DataGrid1.SelBookmarks.Count <> selectlist.Count Then
    '            If selectlist.Count = 1 Then
     '           Retmsg = Retmsg & "this section? Only the record with RecSource='" & myName & "' will be deleted!"
     '           Else
     '            Retmsg = Retmsg & "these sections? Only the record with RecSource='" & myName & "' will be deleted!"
      '
      '          End If
                
     '          Else
      '          Retmsg = Retmsg & "these sections?" '& Me.Adodc1.Recordset!Section_Name & " and etc. ?"
      '          End If
      '          title = "Delete Sections"
     '       End If
      ' End If
                
       ' If Not bAlreadyPopConfirmation Then
       '     If MsgBox(Retmsg, vbYesNo, title) <> vbYes Then
       '       If DataEnvironment2.cnnSection.State = adStateOpen Then
       '         DataEnvironment2.cnnSection.Close
        '      End If
       '         Me.DataGrid1.Visible = True
        '        Screen.MousePointer = vbDefault
        '        DataGrid1_ColResize 0, 0
    
      '      Exit Sub
       '     End If
      '      bAlreadyPopConfirmation = True
       ' End If
        
       '  Me.DataGrid1.Visible = False
         
         
        On Error GoTo ErrorHandling
        
        Dim nozeroMsg As String
        nozeroMsg = "Are you sure you want to delete the selected record(s) in the Section_tbl table?"
    
       'Need to replace the following statment with updateDelete
       If iDaysUntilDelete = 0 Then
       
           If MsgBox(msg, vbYesNo, "Confirmation") = vbYes Then
          
                deleteDocSectionsByPartRecord secName
                deleteDocSectionsByItemRecord secName
                DataEnvironment2.cmdDeleteSection secName
                bDeleteCurrent = True
            Else
                bDeleteCurrent = False
                If DataEnvironment2.cnnSection.State = adStateOpen Then
                    DataEnvironment2.cnnSection.Close
                End If
            End If
            

       Else
       If Not bAlreadyPopConfirmation Then
            If MsgBox(nozeroMsg, vbYesNo, "Confirmation") = vbYes Then
                bDeleteCurrent = True
                bAlreadyPopConfirmation = True
                DataEnvironment2.cmdUpdateDelete CStr(Now), CStr(Now), secName
            Else
                bDeleteCurrent = False
                If DataEnvironment2.cnnSection.State = adStateOpen Then
                    DataEnvironment2.cnnSection.Close
                End If

                Exit For
            End If
            
           
       Else
           If bDeleteCurrent Then
                DataEnvironment2.cmdUpdateDelete CStr(Now), CStr(Now), secName
            End If
       End If
       
    End If
    
        
        
        '''' Comment out the following stuff
        ''Me.Adodc1.Recordset.Requery
        ''Me.Adodc1.Refresh
     '   If Me.Adodc1.Recordset.RecordCount > 0 Then
     '       Me.Adodc1.Recordset.MoveLast
    '    End If
        cmdSort_Click
       If bDeleteCurrent Then
           'currentSelectIndex = CInt(AnyArray(i - 1))
           offset = offset + 1
       End If
         '   currentSelectIndex = CInt(AnyArray(i - 1))
        'End If
        
            If currentSelectIndex > 0 Then
               If Not Me.Adodc1.Recordset.EOF Then
               ' If bDeleteCurrent Then
                '   Me.Adodc1.Recordset.Move (currentSelectIndex - 1)
               ' Else
               ''    Me.Adodc1.Recordset.Move (CInt(AnyArray(i)))
               'End If
                
               Else
                Me.cmdDelete1.Enabled = False
               End If
            Else
                cmdDelete1.Enabled = False
            End If

   ' End If
 '   offset = offset + 1
    Me.DataGrid1.Visible = True
    Next
    Me.DataGrid1.Visible = True
    Screen.MousePointer = vbDefault
    DataGrid1_ColResize 0, 0
    If iDaysUntilDelete > 0 And offset > 0 Then
      MsgBox "Records have been marked for deletion and will be removed based on your DaysUntilDelete setting"
    End If
    If currentDeleteIndex >= 1 Then
     If currentDeleteIndex < Me.Adodc1.Recordset.RecordCount Then
    Me.Adodc1.Recordset.Move currentDeleteIndex, 1
    Else
    Me.Adodc1.Recordset.Move Me.Adodc1.Recordset.RecordCount - 1, 1
    End If
    End If
    
    
    Exit Sub
    
    
ErrorHandling:
MsgBox Err.Description
  If DataEnvironment2.cnnSection.State = adStateOpen Then
            DataEnvironment2.cnnSection.Close
 End If
 Me.DataGrid1.Visible = True
 Screen.MousePointer = vbDefault
 DataGrid1_ColResize 0, 0
     If currentDeleteIndex >= 1 Then
    Me.Adodc1.Recordset.Move currentDeleteIndex, 1
    End If

 
End Sub

Private Function getconfirmationMsg(tablename As String, sec As String) As String
Dim mySelCmd As Command
Dim myret As Recordset
Dim myConn As Connection
Dim msg As String
Dim counter As Integer
Set mySelCmd = New Command
mySelCmd.CommandText = "Select count(*) from " & tablename & " where SOWSection='" & sec & "'"
Set myConn = DataEnvironment2.cnnSection
mySelCmd.ActiveConnection = myConn

msg = ""
On Error GoTo Handling
Set myret = mySelCmd.Execute
counter = myret(0).value
If counter > 0 Then
  If counter = 1 Then
   msg = CStr(counter) & " record from " & tablename & " table"
 Else
   msg = CStr(counter) & " records from " & tablename & " table"
 End If
End If

getconfirmationMsg = msg
Exit Function
Handling:
getconfirmationMsg = msg


End Function

Private Sub deleteDocSectionsByItemRecord(sectionname As String)
Dim myCmd As Command
Dim myConn As Connection
Set myCmd = New Command
myCmd.CommandText = "Delete * from DocSectionsByItem where SOWSection='" & sectionname & "'"
Set myConn = DataEnvironment2.cnnSection
myCmd.ActiveConnection = myConn
On Error GoTo Handling

'If myCmd.ActiveConnection.State <> adStateOpen Then
  '  getpwdConn myCmd.ActiveConnection
   ' myCmd.ActiveConnection.Open
'End If
myCmd.Execute

Handling:
 '   If myCmd.ActiveConnection.State = adStateOpen Then
'        myCmd.ActiveConnection.Close
'End If

End Sub

Private Sub deleteDocSectionsByPartRecord(sectionname As String)
Dim myCmd As Command
Dim myConn As Connection
Set myCmd = New Command
myCmd.CommandText = "Delete * from DocSectionsByPart where SOWSection='" & sectionname & "'"
Set myConn = DataEnvironment2.cnnSection
myCmd.ActiveConnection = myConn
On Error GoTo Handling

'If myCmd.ActiveConnection.State <> adStateOpen Then
    'getpwdConn myCmd.ActiveConnection
   ' myCmd.ActiveConnection.Open
'End If
myCmd.Execute

Handling:
 '   If myCmd.ActiveConnection.State = adStateOpen Then
'myCmd.ActiveConnection.Close
'End If

End Sub

Private Sub moveRecordSet(index As Integer)
Dim i As Integer
 'Me.Adodc1.Recordset.MoveFirst
 'For i = 0 To index - 1
 Me.Adodc1.Recordset.Move index, 1
 'Next
End Sub

   Private Sub MoveRow(ByVal sourceIndex As Integer, ByVal destIndex As Integer)

        Dim items1 As Double
        Dim items2 As Double
        Dim items3 As Double
        Dim tmpRec As Recordset
        items3 = 0
        If sourceIndex = destIndex Then
            Return
        End If
        If (sourceIndex > destIndex) Then
            items1 = getOrderValue(sourceIndex - 1)
            items2 = getOrderValue(destIndex - 1)
            If destIndex >= 2 Then
            items3 = getOrderValue(destIndex - 2)
            Else
            items3 = -1
            End If
            
            If items3 <> -1 Then
                setOrderValue (items2 + items3) / 2, sourceIndex
            Else
                 setOrderValue items2 - 0.1, sourceIndex
            End If


        Else

            items1 = getOrderValue(sourceIndex - 1)
            items2 = getOrderValue(destIndex - 1)
            items3 = getOrderValue(destIndex)
            If items3 <> -1 Then
                setOrderValue (items2 + items3) / 2, sourceIndex
            Else
                 setOrderValue items2 + 0.5, sourceIndex
            End If
        End If
       
     End Sub


Private Sub cmdSort_Click()
   renumber
        
    Me.Adodc1.Refresh
    DataGrid1_ColResize 0, 0
    
End Sub


Private Sub renumber()
  Dim i As Integer
    Dim seclist As New Collection
    If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
    End If

    For i = 1 To seclist.Count
        seclist.Remove 0
    Next i
    
    DataEnvironment2.cmdGetOrderInfo
    
    While Not DataEnvironment2.rscmdGetOrderInfo.EOF
        Dim tmp As String
        tmp = DataEnvironment2.rscmdGetOrderInfo!Section_Name
        seclist.Add tmp
        DataEnvironment2.rscmdGetOrderInfo.MoveNext
    Wend
    
    For i = 1 To seclist.Count
        DataEnvironment2.cmdUpdateOrderInfo i, CStr(seclist.item(i))
    Next i
    
    If DataEnvironment2.cnnSection.State = adStateOpen Then
        DataEnvironment2.cnnSection.Close
    End If
End Sub


Private Sub showProperties(bshowLimit As Boolean, Optional ByVal bAdd As Boolean = False)
    Dim dlgUpdate As New frmUpdate
    
    dlgUpdate.loadInfo Me.Adodc1.Recordset, 0, bAdd
    dlgUpdate.showLimit bshowLimit
        
    If bAdd Then
     dlgUpdate.setAddExternal CDbl(currentSelectIndex + 0.5)
    End If
    dlgUpdate.Show vbModal
    
    If dlgUpdate.isCanceled Then
        DataGrid1_ColResize 0, 0
        Set myfrm = dlgUpdate
        bUpdateDocOK = False
       Exit Sub
    ElseIf dlgUpdate.bOKisClicked Then
    bUpdateDocOK = True
    End If
    'Me.Adodc1.Recordset.MoveFirst
    Set myfrm = dlgUpdate

End Sub

Private Sub cmdProperties_Click()
On Error GoTo Final
    bUpdateDocOK = False
    Dim originalIndex As Integer
    Dim originalIndex2 As Integer
    originalIndex = Me.Adodc1.Recordset.Bookmark
    originalIndex2 = Me.DataGrid1.Row
    showProperties False
    If bUpdateDocOK Then
    bUpdateDocOK = False
    Dim cursor As Integer
    On Error GoTo Final
  '  Me.DataGrid1.Enabled = False
    Me.Adodc1.Recordset.Requery
    Me.Adodc1.Refresh
    DataGrid1_ColResize 0, 0
    Me.DataGrid1.Row = originalIndex2
    Me.Adodc1.Recordset.Move originalIndex - 1, 1
   ' Me.DataGrid1.Enabled = True

End If
Exit Sub
Final:

End Sub

Private Sub viewDoc()
    Dim buff() As Byte
    Dim fh As Integer
    Dim tmpstr As String
    fh = FreeFile()
    buff = Me.Adodc1.Recordset!Word_Doc
    tmpstr = CreateTempFile()
    Open tmpstr For Binary As fh
        Put fh, , buff
    Close fh
    
    openWordDoc myWord1, tmpstr
    
    'OLE1.CreateLink tmpstr
    'OLE1.Visible = True
    DataGrid1_ColResize 0, 0
End Sub


Private Sub cmdViewDoc_Click()
'viewDoc
cmdProperties_Click
End Sub

Private Sub DataGrid1_BeforeUpdate(Cancel As Integer)
    bUpdate = True
End Sub





Private Sub synch()
 '   Me.cmdDelete1.Enabled = Not Me.cmbDocName.Text = ""
   ' Me.cmdGet.Enabled = Not Me.cmbDocName.Text = ""
   ' Me.cmdSave.Enabled = Not Me.txtName.Text = ""
    If DataEnvironment2.cnnSection.State <> adStateOpen Then
      getpwdConn DataEnvironment2.cnnSection
      DataEnvironment2.cnnSection.Open
    End If
    
    Me.cmdDelete1.Enabled = Me.DataGrid1.SelBookmarks.Count <> 0
End Sub

Private Function getDaysUntilDelete() As Integer
    Dim tmp As Variant
    Dim Key As String
On Error GoTo handle


    Key = "Software\VB and VBA Program Settings\CorsPro\Settings"
    tmp = GetRegistryValue(HKEY_CURRENT_USER, Key, "DaysUntilDelete")
    getDaysUntilDelete = CInt(tmp)
    Exit Function
handle:
    getDaysUntilDelete = 0
    
End Function

Private Sub CalculateCounter(brunCount As Boolean)
    Dim runCount As Integer
    Dim tmp As Variant
    Dim Key As String
    Key = "SOFTWARE\Pacific\PG2"

    tmp = GetRegistryValue(HKEY_CLASSES_ROOT, Key, "PP")
    If Not IsEmpty(tmp) Then
        runCount = CInt(tmp)
        
        If runCount > MY_FUN Then
            MsgBox "You run to the limit (" & MY_FUN & " times)" & ",Please contact Elleson Consultant Corp. at 703 834 0836", vbOKOnly, "Warning"
            End
            
        End If
        SetRegistryValue HKEY_CLASSES_ROOT, Key, "PP", runCount + 1
        
    Else
        CreateRegistryKey HKEY_CLASSES_ROOT, Key
        SetRegistryValue HKEY_CLASSES_ROOT, Key, "PP", 0
        
    End If

  '  Me.lbCount.Caption = CStr(runCount)

End Sub


Private Function isDeleteInMix() As Boolean
On Error GoTo handle
    Dim i As Integer
    Dim bDelete As Boolean
    bDelete = False
    
    For i = 0 To Me.DataGrid1.SelBookmarks.Count - 1
            'Me.Adodc1.Recordset.MoveFirst
            'Me.Adodc1.Recordset.Move (Me.DataGrid1.SelBookmarks.item(i) - 1), 1
            If isSameRecSource(Me.Adodc1.Recordset!Section_Name) Then
               bDelete = True
            End If
     Next i
    isDeleteInMix = bDelete
    Exit Function
handle:
   isDeleteInMix = False
End Function

Private Sub DataGrid1_Click()
On Error GoTo L1
currentBookmark = Me.Adodc1.Recordset.Bookmark
'Me.lbRow.Caption = CStr(currentBookmark) & " (click)"
currentLocalIndex = Me.DataGrid1.Row
currentLocalBeginIndex = Me.DataGrid1.FirstRow

If Me.DataGrid1.Row >= 0 Then
currentSelectIndex = Me.DataGrid1.FirstRow + Me.DataGrid1.Row

End If
'Me.cmdDelete1.Enabled = Me.DataGrid1.SelBookmarks.Count <> 0

'If Me.DataGrid1.SelBookmarks.Count > 0 Then
Me.cmdDelete1.Enabled = isDeleteInMix()
'End If

Me.btnUp.Enabled = currentSelectIndex <> 1
Me.cmdDown.Enabled = True ' currentSelectIndex <> Me.Adodc1.Recordset.RecordCount
'Me.lbName.Caption = getSecName
L1:

End Sub

Private Sub DataGrid1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
On Error GoTo Error
 DataGrid1.Columns.item(1).Width = 0
 DataGrid1.Columns.item(5).Width = 0
 DataGrid1.Columns.item(6).Width = 0
 Me.btnUp.Enabled = False
 Exit Sub
Error:
 
End Sub

Private Sub DataGrid1_DblClick()
    cmdProperties_Click
End Sub
Private Function getPreviousSecName() As String
On Error GoTo handle:
Me.Adodc1.Recordset.MovePrevious
getPreviousSecName = Me.Adodc1.Recordset!Section_Name
Me.Adodc1.Recordset.MoveNext
Exit Function
handle:
 getPreviousSecName = ""
End Function

Private Function getNextSecName() As String
On Error GoTo handle:
Me.Adodc1.Recordset.MoveNext
getNextSecName = Me.Adodc1.Recordset!Section_Name
Me.Adodc1.Recordset.MovePrevious
Exit Function
handle:
getNextSecName = ""
End Function
Private Function getSecName() As String
getSecName = Me.Adodc1.Recordset!Section_Name
 
End Function
Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim currentBookmark As Integer
'Dim rowindex As Integer
'currentBookmark = Me.Adodc1.Recordset.Bookmark
'Me.lbRow.Caption = CStr(currentBookmark) & " (KeyDown)"
'Me.lbName.Caption = getSecName

DataGrid1_Click
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Me.Adodc1.Recordset.StayInSync = True
Dim currentBookmark As Integer
On Error GoTo handle:
currentBookmark = Me.Adodc1.Recordset.Bookmark
Exit Sub
'Me.lbtemp.Caption = CStr(currentBookmark)
'DataGrid1.SetFocus
handle:

End Sub

Private Sub DataGrid1_Scroll(Cancel As Integer)
Cancel = False
Dim currentBookmark As Integer
currentBookmark = Me.Adodc1.Recordset.Bookmark
Me.lbmm.Caption = CStr(Me.DataGrid1.FirstRow)
'Me.lbRow.Caption = CStr(currentBookmark) & " (Scroll)"
'Me.lbName.Caption = getSecName

End Sub

Private Sub DataGrid1_SelChange(Cancel As Integer)
'Dim currentBookmark As Integer
'currentBookmark = Me.Adodc1.Recordset.Bookmark
'Me.lbmm.Caption = CStr(currentBookmark) & " (SelChange)" & " Secanme=" & getSecName
'Me.lbName.Caption = getSecName


End Sub

Private Sub Form_Activate()
    Dim Directory_Root As String
    Directory_Root = GetRegValue(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\CorsPro\Settings", "Directory_Company")

    'app.HelpFile = Directory_Root & "PropGen\LibrMgr.chm"
    
   
    'Me.HelpContextID = ID_SECTIONMGR
    'Me.HelpContextID = 1000

    resetConnectStr

    Me.Adodc1.Refresh
    'Me.cmdDelete1.Enabled = Me.Adodc1.Recordset.RecordCount > 0
    DataGrid1_ColResize 0, 0
    Me.cmdDelete1.Enabled = Me.DataGrid1.SelBookmarks.Count <> 0
    bSectionActive = True
    
     myName = getClientName()
     If Me.Adodc1.Recordset.RecordCount > 0 Then
    Me.cmdDelete1.Enabled = Me.Adodc1.Recordset!RecSource = myName
    End If
   defaultDocType = getDefaultDocType()
   DataGrid1.Columns(0).Locked = True
   
   
   iDaysUntilDelete = getDaysUntilDelete()
   'Me.btnUp.Enabled = True
   currentSelectIndex = 1
   
   TrackIndex = 1
   
End Sub

Public Sub setParent(parent As frmLibManager)
Set myParent = parent
End Sub

Public Sub setConnStr()
 resetConnectStr
End Sub


Private Sub resetConnectStr()
 Dim myStr As String
 myStr = Me.Adodc1.ConnectionString
 Dim dbFullpath As String
 If Not bconvertAdo Then
 
     If DataEnvironment2.cnnSection.State = adStateOpen Then
     DataEnvironment2.cnnSection.Close
     End If
     
     

    dbFullpath = choseDB

    Me.Adodc1.ConnectionString = getpwdconnStr(replaceDataBase(myStr, "Data Source=" & dbFullpath))
    'Late bind is very important, otherwise we cannot make it work if connection string misses valid Data Source
    Set Me.DataGrid1.DataSource = Me.Adodc1
    Debug.Print "Col no2=" & CStr(DataGrid1.Columns.Count)
    DataGrid1_ColResize 0, 0
    DataGrid1.Refresh
    
    bconvertAdo = True
 End If

End Sub


Private Sub Form_Deactivate()
  bconvertAdo = False
  bSectionActive = False
   If Not bTempUnload Then
     bTempActive = True
   End If
End Sub

Private Sub Form_GotFocus()
DataGrid1_ColResize 0, 0
bTempActive = False
bSectionActive = True

End Sub

Private Sub Form_Load()
diff_H = 1230
diff_W = 435
 'diff_H = Me.Height - DataGrid1.Height
 'diff_W = Me.Width - DataGrid1.Width
 o_H = Me.Height
 o_W = Me.Width
 diff_BL = Me.cmdDelete1.Left - Me.Left
 old_H = Me.Height
 bSectionUnload = False
 SectionFileLen = 30
 bForcedByMDIP = False
 renumber

End Sub


Public Sub shutdownByParent()
bForcedByMDIP = True
End Sub

Private Sub Form_LostFocus()
bSectionActive = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

 If bForcedByMDIP Then
 Exit Sub
 End If


 If MsgBox("Do you want to close Section Manager?", vbYesNo, "Confirmation") = vbYes Then
   
   Else
   Cancel = 1
 End If

End Sub

Private Sub addDebug(msg As String)
'Text1.Text = Text1.Text & msg & vbCrLf
End Sub
Private Sub Form_Resize()
  'DataGrid1.Move o_top, o_left, Me.Width - 50, Me.Height - 50
  
  ''If o_W > Me.Width Then
   ' Me.Width = o_W
  'End If
  'If o_H > Me.Height Then
  '  Me.Height = o_H
 'End If
    'Dim diffH As Integer
    'diffH = Me.Height - old_H
    'Me.cmdAdd.Top = Me.cmdAdd.Top + diffH - 10
    'Me.cmdProperties.Top = Me.cmdProperties.Top + diffH - 10
    'Me.cmdViewDoc.Top = Me.cmdViewDoc.Top + diffH - 10
    'Me.cmdDelete1.Top = Me.cmdDelete1.Top + diffH - 10
  
  If Me.Width > diff_W Then
    DataGrid1.Width = Me.Width - diff_W
   
  End If
  If Me.Height > diff_H Then
    DataGrid1.Height = Me.Height - diff_H
  End If
  
  old_H = Me.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmdSort_Click
bconvertAdo = False
bSectionActive = False

bSectionUnload = True
If Not bTempUnload Then
  bTempActive = True
End If

End Sub



Private Sub OLE1_Click()
'    ViewDoc
End Sub

Private Sub txtName_Change()
    synch
End Sub




