VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUpdate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Update Dialog"
   ClientHeight    =   3120
   ClientLeft      =   30
   ClientTop       =   240
   ClientWidth     =   5910
   ControlBox      =   0   'False
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5280
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxtList 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   22
      Text            =   "frmUpdate.frx":1CCA
      Top             =   3360
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   1320
   End
   Begin VB.ComboBox cmbDocType 
      Height          =   315
      ItemData        =   "frmUpdate.frx":1CD0
      Left            =   1320
      List            =   "frmUpdate.frx":1CD2
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   990
      Width           =   3375
   End
   Begin VB.CheckBox ckMultiple 
      Caption         =   "Multiple Documents"
      Height          =   372
      Left            =   480
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.ComboBox cmbKeepStyle 
      Height          =   315
      ItemData        =   "frmUpdate.frx":1CD4
      Left            =   0
      List            =   "frmUpdate.frx":1CDE
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   3372
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   2475
      Width           =   852
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "OK"
      Height          =   255
      Left            =   4080
      TabIndex        =   13
      Top             =   2475
      Width           =   852
   End
   Begin VB.ComboBox cmbType 
      DataField       =   "Object_Type"
      Height          =   315
      ItemData        =   "frmUpdate.frx":1CEB
      Left            =   1320
      List            =   "frmUpdate.frx":1CF5
      TabIndex        =   12
      Top             =   540
      Width           =   3372
   End
   Begin VB.CommandButton cmbSelectDoc 
      Caption         =   "..."
      Height          =   372
      Left            =   5040
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.TextBox txtDocName 
      Enabled         =   0   'False
      Height          =   288
      Left            =   0
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   3612
   End
   Begin VB.TextBox txtDescription 
      Height          =   888
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1440
      Width           =   4425
   End
   Begin VB.TextBox txtKeep_Style 
      DataField       =   "Keep_Style"
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.TextBox txtOrder_Number 
      DataField       =   "Order_Number"
      Height          =   285
      Left            =   -120
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.TextBox txtSection_Name 
      DataField       =   "Section_Name"
      Height          =   285
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   1
      Top             =   120
      Width           =   3372
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View/Edit"
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   2475
      Width           =   1092
   End
   Begin VB.Label lbLimit 
      Caption         =   "(<=30 chars)"
      Height          =   285
      Left            =   4800
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lbWriteTime 
      Height          =   375
      Left            =   6840
      TabIndex        =   20
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lbreateTime 
      Height          =   255
      Left            =   6840
      TabIndex        =   19
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Document"
      Height          =   285
      Left            =   45
      TabIndex        =   9
      Top             =   2282
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Doc Type:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1050
      Width           =   750
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Object_Type:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   570
      Width           =   1095
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Order_Number:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblFieldLabel 
      AutoSize        =   -1  'True
      Caption         =   "Section_Name:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2310
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Temp holders
Dim records As Recordset
Dim bLoadNewDoc As Boolean
Dim docName As String
Dim docBuff() As Byte
Dim bInsertRecord As Boolean

Dim currentWorkingFile As String
'Place holders for current value
'SectionName
Dim currentSec As String
'Order_Number
Dim currentOrder As Double
'Keep_Style
Dim currentStyle As String
'DocType
Dim currentDocType As String
'Description
Dim currentDesc As String
'Type
Dim currentType As String
Dim bSecNameChanged As Boolean
Dim OriginalSecName As String
Dim bAddExternal As Boolean
Dim charlist() As Byte
Dim myORange As Object 'Range

'Current active document changed flag
Dim bDocChanged As Boolean
Dim bWorkDocChanged As Boolean

Dim bStartCheck As Boolean
Dim bBeingClosed As Boolean
Dim bOpenDoc As Boolean
Dim bMultipleDoc As Boolean
'Original word doc buff
Dim o_buff() As Byte
Dim o_filelength As Long
Dim orText, lastText

Dim buff() As Byte
'Indicate whether to enable Add button or not
Dim bAddOK As Boolean

Dim bIsCanceled As Boolean
Dim lastIndex As Integer
Dim bWorkdDocChanged As Boolean
Dim myName As String
Public bIsTerminated As Boolean
Dim originalHeader As String
Dim originalFooter As String
Dim startStamp As Date
Dim closeStamp As Date
Dim bCheckTime As Boolean

'Used for tap word application event
'Private WithEvents myWord1 As Word.Application
Private myWord1 As Object
Public bOKisClicked As Boolean

'check the timestamp of the tmp file, if there is no change in the
'modified date, there no action to save the file back to database
'Timestamp before modified
Dim bTimestamp As Date
'Timestamp after modified
Dim aTimestamp As Date

Public Function isAddedOK() As Boolean

    isAddedOK = bAddOK
End Function
Public Sub loadInfo(rec As Recordset, Optional currentIndex As Double, Optional bNew As Boolean)
    If bNew Then
        Me.txtSection_Name.Text = ""
        Me.txtOrder_Number.Text = CStr(Abs(currentIndex - 0.5))
        Me.txtDescription.Text = ""
        Me.cmbType.Text = "Internal"
        Me.cmbKeepStyle.Text = "Yes"
        Me.cmbDocType.Text = defaultDocType
        bInsertRecord = True
        currentStyle = Me.cmbKeepStyle.Text
        currentDocType = Me.cmbDocType.Text
        currentDesc = Me.txtDescription.Text
        currentType = Me.cmbType.Text
        currentSec = ""
        syn
        bAddOK = False
        Exit Sub
    End If
    Me.txtSection_Name.Enabled = bNew
    On Error Resume Next
    Me.txtSection_Name.Text = rec!Section_Name
    OriginalSecName = rec!Section_Name
    Me.cmbDocType.Text = rec!doctype
    Me.txtOrder_Number.Text = CStr(rec!Order_Number)
    Me.cmbKeepStyle.Text = rec!Keep_Style
    Me.txtDescription.Text = rec!Description
    Me.cmbType.Text = rec!Object_Type
    
    If txtOrder_Number.Text <> "" Then
        currentOrder = CDbl(txtOrder_Number.Text)
    Else
        currentOrder = -1
    End If
    currentStyle = Me.cmbKeepStyle.Text
    currentDocType = Me.cmbDocType.Text
    currentDesc = Me.txtDescription.Text
    currentType = Me.cmbType.Text
    currentSec = Me.txtSection_Name.Text

    currentWorkingFile = Me.txtSection_Name.Text
    
    Set records = rec
    syn
    bLoadNewDoc = False
    
    bStartCheck = True
    
End Sub


Private Sub ckMultiple_Click()
   bMultipleDoc = ckMultiple.value
End Sub

Private Sub cmbDocType_Click()
syn
End Sub

Private Sub cmbKeepStyle_Click()
syn
End Sub

Public Sub setAddExternal(totalrecord As Double)
bAddExternal = True
 Me.cmbType.Text = "External"
 Me.cmbType.Enabled = False
 Me.cmdView.Enabled = False
 Me.txtOrder_Number = CStr(totalrecord)
 Me.txtSection_Name = ""
 bInsertRecord = True
 syn
End Sub

Private Sub cmbSelectDoc_Click()
    CommonDialog1.ShowOpen
    docName = CommonDialog1.filename
    
    Me.cmdUpdate.Enabled = True
    
    
    If docName = "" Then
     Exit Sub
    End If
        
    buff = OpenToDBNew(docName)
    bLoadNewDoc = True
    If docName <> "" Then
       Me.txtSection_Name = Mid(docName, InStrRev(docName, "\") + 1, InStrRev(docName, ".") - InStrRev(docName, "\") - 1)
    End If
    Me.Caption = Me.txtSection_Name & "'s Properties"
    OriginalSecName = Me.txtSection_Name
    syn
    
End Sub
Public Function getDocType() As String
getDocType = Me.cmbDocType.Text
End Function
Public Function getOrderNumber() As Double
   If Me.txtOrder_Number <> "" Then
   getOrderNumber = CDbl(Me.txtOrder_Number)
   Else
   getOrderNumber = 0
   End If
End Function
Public Function setOrderNumber(no As Integer)
   Me.txtOrder_Number.Text = CStr(no)
End Function

Public Function getKeepStyle() As String
getKeepStyle = Me.cmbKeepStyle.Text
End Function
Public Function getDesc() As String
   getDesc = Me.txtDescription.Text
End Function
Public Function getSecName() As String
 getSecName = Me.txtSection_Name.Text
 
End Function

Private Sub cmbType_Click()
    syn
End Sub


Private Sub UpdateRec()
    Dim i As Integer
      Dim seclist As New Collection
      
      myName = getClientName()
  
      If currentWorkingFile = "" Or InStr(1, currentWorkingFile, ".tmp") = 0 Then
          ' If records!Word_doc Is Nothing Then
            docBuff = loadDoc(getSecName()) 'records!Word_doc
         '   End If
       
       End If
       
   
      If bLoadNewDoc Then
        docBuff = buff
        bLoadNewDoc = False
      ElseIf currentWorkingFile <> "" And InStr(1, currentWorkingFile, ".tmp") <> 0 Then
        Dim fh As Integer
        Dim flen As Long
        Dim dd As String
        currentWorkingFile = formatDocxToDoc(currentWorkingFile)
        addLog "Reopen the working file from " & currentWorkingFile
        fh = FreeFile()
        addLog "Save the current updated file into DB: " & currentWorkingFile
        'flen = FileLen(currentWorkingFile) + 1
        flen = FileLen(currentWorkingFile) ' + 1
        addLog "File Size= " & flen

        ReDim buff(flen - 1) As Byte
        'buff = records!Word_Doc
        dd = currentWorkingFile
        Open currentWorkingFile For Binary As #fh
        Get #fh, 1, buff
        Close fh
        docBuff = buff
        On Error Resume Next
        'myWord1.ActiveDocument.Close
        If Not myWord1 Is Nothing Then
            myWord1.Documents(getPureName(currentWorkingFile)).Close
        End If
       currentWorkingFile = ""
      End If
      'MsgBox "Will Update " & dd, vbOKOnly
      
      On Error GoTo Error:
    
    ' Get ClientName
      If Me.cmbType.Text = "Internal" Then
      
      addLog "Update Record for internal document"


        If Not isSameRecSource(Me.txtSection_Name.Text) Then
        
            If DataEnvironment2.cnnSection.State <> adStateOpen Then
              getpwdConn DataEnvironment2.cnnSection
              DataEnvironment2.cnnSection.Open
            End If

            If bWorkDocChanged Or bDocChanged Then
                addLog "(11)Update " & records!Section_Name
                addLog " With the following data:"
                addLog " OrderNumber:" & CDbl(Me.txtOrder_Number.Text)
                addLog " Type:" & Me.cmbType.Text
                addLog " DocType:" & Me.cmbDocType.Text
                addLog " cmbKeepStyle:" & Me.cmbKeepStyle.Text
                addLog " Decription: " & Me.txtDescription.Text
                addLog " Name:" & myName
                addLog " Time/Date:" & CDate(Now)
                DataEnvironment2.cmdUpdateNoMyRec CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbDocType.Text, "yes", Me.txtDescription.Text, docBuff, myName, CDate(Now), records!Section_Name
            Else
                addLog "(12)Update " & records!Section_Name

                DataEnvironment2.UpdatedNoMyRecNoWordDocChange CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbDocType.Text, "yes", Me.txtDescription.Text, records!Section_Name
            End If
            
        Else
           If DataEnvironment2.cnnSection.State <> adStateOpen Then
            getpwdConn DataEnvironment2.cnnSection
            DataEnvironment2.cnnSection.Open
           End If

           If bWorkDocChanged Or bDocChanged Then
                addLog "(21)Update " & records!Section_Name
                addLog " With the following data:"
                addLog " OrderNumber:" & CDbl(Me.txtOrder_Number.Text)
                addLog " Type:" & Me.cmbType.Text
                addLog " DocType:" & Me.cmbDocType.Text
                addLog " cmbKeepStyle:" & Me.cmbKeepStyle.Text
                addLog " Decription: " & Me.txtDescription.Text
                addLog " Name:" & myName
                addLog " Time/Date:" & CDate(Now)
             DataEnvironment2.UpdateMyOwnRec CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbDocType.Text, "yes", Me.txtDescription.Text, docBuff, CDate(Now), records!Section_Name
            Else
            
             addLog "(22)Update my" & records!Section_Name

             DataEnvironment2.UpdateMyOwnRecNoWordDocChange CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbDocType.Text, "yes", Me.txtDescription.Text, records!Section_Name
            End If
            
        End If
        
       
            
      Else
      
      If Not isSameRecSource(Me.txtSection_Name.Text) Then
           ' If bWorkDocChanged Then
           '  DataEnvironment2.cmdUpdateNoMyRec CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbKeepStyle.Text, Me.txtDescription.Text, Empty, ClientName, CDate(Now), records!Section_Name
           ' Else
          If DataEnvironment2.cnnSection.State <> adStateOpen Then
            getpwdConn DataEnvironment2.cnnSection
            DataEnvironment2.cnnSection.Open
           End If

             DataEnvironment2.UpdatedNoMyRecNoWordDocChange CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbDocType.Text, Me.cmbKeepStyle.Text, Me.txtDescription.Text, records!Section_Name
          '  End If

       ' DataEnvironment2.cmdUpdateNoMyRec CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbKeepStyle.Text, Me.txtDescription.Text, Empty, ClientName, CDate(Now), records!Section_Name
      Else
        'DataEnvironment2.UpdateMyOwnRec CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbKeepStyle.Text, Me.txtDescription.Text, Empty, CDate(Now), records!Section_Name
            ' If bWorkDocChanged Then
             'DataEnvironment2.UpdateMyOwnRec CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbKeepStyle.Text, Me.txtDescription.Text, docBuff, CDate(Now), records!Section_Name
           ' Else
            If DataEnvironment2.cnnSection.State <> adStateOpen Then
              getpwdConn DataEnvironment2.cnnSection
              DataEnvironment2.cnnSection.Open
            End If

             DataEnvironment2.UpdateMyOwnRecNoWordDocChange CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbDocType.Text, Me.cmbKeepStyle.Text, Me.txtDescription.Text, records!Section_Name
          '  End If
   
      End If
      End If
     
      If DataEnvironment2.cnnSection.State = adStateOpen Then
        DataEnvironment2.cnnSection.Close
      End If
      bWorkDocChanged = False
      Me.Hide
      Exit Sub
      
Error:
     If DataEnvironment2.cnnSection.State = adStateOpen Then
      DataEnvironment2.cnnSection.Close
      End If
      MsgBox Err.Description, , "Warning"
End Sub


Public Sub AddRec()
    Dim i As Integer
      Dim seclist As New Collection
      If DataEnvironment2.cnnSection.State <> adStateOpen Then
      getpwdConn DataEnvironment2.cnnSection
      DataEnvironment2.cnnSection.Open
      End If
      If bLoadNewDoc Then
      docBuff = buff
      bLoadNewDoc = False
      Else
      ' If Not records Is Nothing Then
       docBuff = loadDoc(getSecName()) 'records!Word_doc
                
      ' End If
      End If
      
      On Error GoTo Error
      If Me.cmbType.Text = "Internal" Then
 
        DataEnvironment2.cmdInsertDoc Me.txtSection_Name, CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbDocType.Text, docBuff, Me.cmbKeepStyle.Text, Me.txtDescription.Text, myName, Now
      Else
        DataEnvironment2.cmdInsertDoc Me.txtSection_Name, CDbl(Me.txtOrder_Number.Text), Me.cmbType.Text, Me.cmbDocType.Text, Empty, Me.cmbKeepStyle.Text, Me.txtDescription.Text, myName, Now
      
      End If
     
      If DataEnvironment2.cnnSection.State = adStateOpen Then
      DataEnvironment2.cnnSection.Close
      End If
      Me.Hide
      bAddOK = True
      Exit Sub
      
Error:
      If DataEnvironment2.cnnSection.State = adStateOpen Then
      DataEnvironment2.cnnSection.Close
      End If
      MsgBox Err.Description & "Inform=" & inputMsg, , "Warning"
      bAddOK = False
End Sub

Private Sub cmdEditDoc_Click()
'OLE1_Click
End Sub

Private Sub cmdCancel_Click()
If bAddExternal Then
    bAddExternal = False
    bIsCanceled = True
    Me.Hide
    Exit Sub
End If

 On Error Resume Next
 If Not myWord1 Is Nothing Then
 Dim docName As String
 Dim Doc As Object
 docName = getPureName(currentWorkingFile)
 Set Doc = myWord1.Documents(docName)
 Doc.Save
 Doc.Close
 End If
 Me.Hide
 bIsCanceled = True
End Sub

Public Function isCanceled() As Boolean
    isCanceled = bIsCanceled
End Function

Public Function getPureName(Path As String)
    getPureName = Right(Path, Len(Path) - InStrRev(Path, "\"))
End Function
Private Sub cmdUpdate_Click()

If bAddExternal Then
    bAddExternal = False
    Me.Hide
    Exit Sub
End If
Dim bNumber As Boolean
Dim bKey As Boolean
bOpenDoc = False
'bDocChanged = False
If Me.txtOrder_Number.Text = "" Then
    bNumber = False
Else
    bNumber = True
End If
If Me.txtSection_Name.Text = "" Or Not IsNumeric(txtOrder_Number.Text) Then
   bKey = False
Else
   bKey = True
End If


If Not (bKey And bNumber) Then
 MsgBox "Invalid Section name or Order number", vbOKOnly, "Warning"
 Exit Sub
End If

If isDuplicateOrderNo(CDbl(CDbl(Me.txtOrder_Number.Text))) Then
   MsgBox "Duplicate Order Number! please try again!"
   Me.txtOrder_Number.Text = ""
   Me.txtOrder_Number.SetFocus
   Exit Sub
   
Else
    If bSecNameChanged Then
        bSecNameChanged = False
        AddRec
       'delete current record
       
        If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
        End If
        
        DataEnvironment2.cmdDeleteSection OriginalSecName
             
        If DataEnvironment2.cnnSection.State = adStateOpen Then
        DataEnvironment2.cnnSection.Close
        End If

        
       'add new rec
       
    
    Else
    
    If Me.cmbDocType.Text = "" Then
    MsgBox "DocType cannot be empty string, please select one of the following four values:" & vbCrLf & "'SOW', 'Proposal' , 'Proposal + SOW' and 'Internal'"
    Exit Sub
    End If

        If bInsertRecord Then
            AddRec
        Else
        
            aTimestamp = fileModified(currentWorkingFile)
            If bTimestamp <> aTimestamp Then
                UpdateRec
            End If
            
        End If
        bOKisClicked = True
    End If
    
 End If
 bWorkdDocChanged = False
 
End Sub


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

Private Sub syn()
    If Me.cmbType.Text = "Internal" Then
        Me.txtDocName.Text = "Word document"
        Me.cmbSelectDoc.Enabled = True
        Me.cmdView.Enabled = True
    Else
        Me.cmbSelectDoc.Enabled = False
        Me.txtDocName.Text = "N/A"
        Me.cmdView.Enabled = False
    End If
    If bInsertRecord Then
     If Me.txtSection_Name.Text = "" Then
        Me.Caption = "Add New Section"
     Else
        Me.Caption = Me.txtSection_Name.Text + "'s Properties"
     End If
     Me.cmdUpdate.Caption = "OK"
     'Me.txtSection_Name.Enabled = True
     Me.cmdView.Enabled = bLoadNewDoc
    Else
     
     Me.Caption = Me.txtSection_Name.Text + "'s Properties"
     Me.cmdUpdate.Caption = "OK"
     'Me.txtSection_Name.Enabled = False
    End If
    
 
    


    If Me.txtOrder_Number.Text <> "" And Me.txtSection_Name <> "" Then
    Me.cmdUpdate.Enabled = currentSec <> Me.txtSection_Name.Text Or _
                            currentOrder <> CDbl(Me.txtOrder_Number.Text) Or _
                            currentStyle <> Me.cmbKeepStyle.Text Or _
                            currentDocType <> Me.cmbDocType.Text Or _
                            currentDesc <> Me.txtDescription.Text Or _
                            currentType <> Me.cmbType.Text Or bDocChanged Or _
                            bLoadNewDoc
                            
    Else
    Me.cmdUpdate.Enabled = False
    'currentSec <> Me.txtSection_Name.Text Or _
     '                       currentStyle <> Me.cmbKeepStyle.Text Or _
      '                      currentDesc <> Me.txtDescription.Text Or _
       '                     currentType <> Me.cmbType.Text Or bDocChanged Or _
        '                    bLoadNewDoc
    End If
    
    'Change the logic based upon Brian's request April 24 2007
    Me.cmdUpdate.Enabled = True
    
    If currentSec <> Me.txtSection_Name.Text Then
      
    
    End If
    
    
    
End Sub

Public Function getFileWithDefaultPath(filename As String) As String
On Error Resume Next
If myWord1 Is Nothing Then
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
End If

    getFileWithDefaultPath = filename
End Function

Public Function fileModified(ByVal sFile As String) As Date
   Dim objFSO As New Scripting.FileSystemObject
   If objFSO.FileExists(sFile) Then
       Dim objFile As file
       Set objFile = objFSO.GetFile(sFile)
       fileModified = objFile.DateLastModified
   End If
   Set objFSO = Nothing
End Function
Public Function loadDoc(sname As String) As Byte()
   Dim retf() As Byte
    If DataEnvironment2.cnnSection.State <> adStateOpen Then
    getpwdConn DataEnvironment2.cnnSection
    DataEnvironment2.cnnSection.Open
    End If
    
    On Error Resume Next
    
    
    DataEnvironment2.cmdLoadDoc sname
   ' DataEnvironment2.rscmdLoadDoc.MoveNext
    retf = DataEnvironment2.rscmdLoadDoc!Word_Doc
    
    If DataEnvironment2.cnnSection.State = adStateOpen Then
        DataEnvironment2.cnnSection.Close
    End If
    loadDoc = retf
   
End Function
Private Sub cmdView_Click()
    'Dim buff() As Byte
    Set myWord1 = Nothing
    
    Dim fh As Integer
    Dim fname As String
    fh = FreeFile()
    getFileWithDefaultPath ""
    
    
    
    If Not bLoadNewDoc Then
    If Me.cmbType.Text = "Internal" Then
    On Error GoTo handleError
        'o_buff = records!Word_doc
        o_buff = loadDoc(getSecName())
        
        
addLog "Download doc from DB with size=" & UBound(o_buff) + 1
        'getFileWithDefaultPath ""
        currentWorkingFile = CreateTempFile()
        
addLog "Create Tmp file " & currentWorkingFile
    GetFileTimeInfo currentWorkingFile, , startStamp
        Me.lbreateTime.Caption = CStr(startStamp)
        End If
        
    Else
        currentWorkingFile = docName
addLog "Save file into " & docName

    End If
    
    
    Open currentWorkingFile For Binary As fh
    Put fh, , o_buff
    Close fh
   
    'o_filelength = FileLen(currentWorkingFile)
    'Dim doc As Word.Document

    bTimestamp = fileModified(currentWorkingFile)
    addLog "Current:" & currentWorkingFile
    Set Doc = openWordDocM(currentWorkingFile) ' openWordDoc(myWord1, currentWorkingFile)
    If Doc Is Nothing Then
        bOpenDoc = False
    addLog "Error:Fail to open this document, please check the document format to make sure it's word97 or above"
       ' MsgBox "Fail to open this document, please check the document format to make sure it's word97 or above", vbOKOnly, "Error"
    Exit Sub
    End If
    
    o_filelength = myWord1.ActiveDocument.Characters.Count
    
    'ReDim charlist(o_filelength - 1)
    'Dim d As Long
    
    Set myORange = myWord1.ActiveDocument.Range(0, o_filelength - 1)
    orText = myORange.Text
    o_filelength = myORange.StoryLength
    cmdView.Enabled = False
    bOpenDoc = True
    Dim mySec As Object ' Section
    Set mySec = Doc.Sections.First
    originalHeader = mySec.Headers(1).Range.Text 'wdHeaderFooterPrimary).Range.Text

    originalFooter = mySec.Footers(1).Range.Text 'wdHeaderFooterPrimary).Range.Text
    Me.Timer1.Enabled = True
    bCheckTime = False
     Me.cmdView.Enabled = False
    'OLE1.CreateLink currentWorkingFile
    Exit Sub
    
handleError:
    MsgBox "Cannot open this document"
End Sub



Private Sub Form_Activate()
bOKisClicked = False
bIsCanceled = False
bWorkdDocChanged = False
bCheckTime = False
bTimestamp = CDate("1/1/1990")
aTimestamp = CDate("1/1/1990")

End Sub

Private Sub Form_GotFocus()
syn
End Sub

Private Sub Form_Load()
bOpenDoc = False
bStartCheck = False
'bEnableOKButton = False
bIsCanceled = False
bOKisClicked = False
Set mycol = populateDocType()
For i = 1 To mycol.Count
Me.cmbDocType.AddItem mycol.item(i)
Next
Me.cmbDocType.Text = Me.cmbDocType.List(0)

End Sub

Public Sub setButtonFocus(bAdd As Boolean)
  If bAdd Then
    Me.cmbSelectDoc.TabIndex = 1
  End If
End Sub

Private Sub Form_LostFocus()
bDocChanged = False
syn
End Sub

Private Sub Form_Terminate()
bIsTerminated = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If bAddExternal Then
    bAddExternal = False
    bIsCanceled = True
    Me.Hide
    Exit Sub
End If
End Sub
Private Sub hanleDocumentBeforeClose(ByVal Doc As Object, Cancel As Boolean) 'Word.Document, Cancel As Boolean)
If bDocChanged Then 'myWord1.ActiveDocument.Saved Or
    bDocChanged = True
    bWorkDocChanged = True
    currentWorkingFile = myWord1.ActiveDocument.fullname
    Set myWord1 = Nothing
Else
    bDocChanged = False
    bWorkDocChanged = False
End If

Me.cmdView.Enabled = True
'Me.cmdUpdate.Enabled = True
bBeingClosed = True
'On Error GoTo Final
' If Not myWord1 Is Nothing Then
'  myWord1.Documents(getPureName(myWord1.ActiveDocument.name)).Close
 
' End If
'Final:
 syn
bDocChanged = False
bBeingClosed = False
Me.cmdView.Enabled = True
                 '   Me.Timer1.Enabled = False
                    bCheckTime = False

'myWord1.ActiveDocument.Saved = False
'bWorkdDocChanged = True
End Sub


'Private Sub myWord1_DocumentBeforeClose(ByVal doc As Word.Document, Cancel As Boolean)

'Call hanleDocumentBeforeClose(doc, Cancel)
'End Sub

Private Sub handleDocumentBeforeSave(ByVal Doc As Object, SaveAsUI As Boolean)
 If bBeingClosed Then
 Exit Sub
 End If
  Dim nowDate As Date
  
  Dim locaR As Object ' Range
  Dim lens, k As Long
  

 myWord1.ActiveDocument.Saved = False
  If bOpenDoc Then ' And myWord1.ActiveDocument.Saved Then
       'Doc.AcceptAllRevisionsShown
       If SaveAsUI Then
       'Doc.SaveAs (Doc.FullName)
       Else
       Doc.Save
       End If
    End If
    
          'bDocChanged = True
     '  If MsgBox("Do you want to save this document?", vbYesNo) = vbYes Then
     '      bDocChanged = True
     '   Else
     '    bDocChanged = False
    '   End If
    
    
 ''       currentWorkingFile = doc.FullName

 ''       lens = doc.Characters.Count
        
 ''       If lens <> o_filelength Then
 ''         bDocChanged = True
  ''      Else
 ''          Dim v1, v2
 ''          Set locaR = doc.Range(0, lens - 1)
           'bDocChanged = locaR.IsEqual(myORange)
 ''          lens = Len(locaR.Text)
''           For k = 1 To lens - 1
 ''           v1 = Mid(orText, k, 1)
 ''           v2 = Mid(locaR.Text, k, 1)
          
 ''           If v1 <> v2 Then
''            bDocChanged = True
''            End If
''           Next
''        End If
  
  '      Dim mySection As Section
  '      Dim myHF As HeaderFooter
  '      Set mySection = doc.Sections.First
  '      Dim headerstr As String
  '      Dim footer As String
        
         'For i = 0 To mySection.Headers.Count - 1
  '     headerstr = mySection.Headers(wdHeaderFooterPrimary).Range.Text
 '      footer = mySection.Footers(wdHeaderFooterPrimary).Range.Text
 '
      ' bDocChanged = bDocChanged Or originalHeader  .equals(headerstr) And originalFooter.equals(originalFooter)
    
      'Set myORange = myWord1.ActiveDocument.Range(0, o_filelength - 1)
      
     'Set DocProps = Doc.BuiltInDocumentProperties
      'lens = FileLen(Doc.FullName)
      'bDocChanged = lens <> o_filelength
     ' If lens <> o_filelength Then
    '  bDocChanged = True
    '  Else
      'bDocChanged = IsTheSame(
     ' End If
 ''     syn
   '   bOpenDoc = False
 '' End If
  
  Exit Sub
Handler:
 

End Sub
'Private Sub myWord1_DocumentBeforeSave(ByVal doc As Object, SaveAsUI As Boolean, Cancel As Boolean) ' Word.Document, SaveAsUI As Boolean, Cancel As Boolean)
'Call handleDocumentBeforeSave(doc, SaveAsUI)
  
'End Sub



Private Sub myWord1_DocumentOpen(ByVal Doc As Object) 'Word.Document)
    bFirstOpen = True
End Sub

Private Sub myWord1_Quit()
Dim s
s = 0
End Sub

'Private Sub myWord1_WindowActivate(ByVal doc As Word.Document, ByVal Wn As Word.Window)
'bFirstOpen = True
'End Sub

Private Sub myWord1_WindowActivate(ByVal Doc As Object, ByVal Wn As Object) ' Word.Document, ByVal Wn As Word.Window)
bFirstOpen = True
End Sub


Private Sub Timer1_Timer()
Dim tmpDate As Date
   If Not bCheckTime Then
      GetFileTimeInfo currentWorkingFile, , startStamp
    bCheckTime = True
    End If
    
    
    If GetFileTimeInfo(currentWorkingFile, tmpDate, closeStamp) Then
        Me.lbreateTime.Caption = CStr(startStamp)
        Me.lbWriteTime.Caption = CStr(closeStamp)
    
        If bOpenDoc Then
                 bDocChanged = DateDiff("s", startStamp, closeStamp)
                
                
                If bDocChanged Then
                    bOpenDoc = False
                    Me.cmdUpdate.Enabled = True
                Else
                
                   Me.cmdUpdate.Enabled = False
                   'Change the logic based upon Brian's request April 24 2007
                    Me.cmdUpdate.Enabled = True
                End If
           ' End If
        Me.cmdView.Enabled = True
        End If
    End If
End Sub

Private Sub txtDescription_Change()
syn
End Sub

Private Sub txtKeep_Style_Change()
syn
End Sub

Private Sub txtOrder_Number_KeyUp(KeyCode As Integer, Shift As Integer)
 If Not IsNumeric(Me.txtOrder_Number.Text) Then
  MsgBox "Order number must be numeric !", , "Warning"
  If Me.txtOrder_Number.Text <> "" Then
  Me.txtOrder_Number.Text = Left(Me.txtOrder_Number.Text, Len(Me.txtOrder_Number.Text) - 1)
  Me.txtOrder_Number.SelStart = Len(Me.txtOrder_Number.Text)
  End If
 End If
 syn
End Sub


Private Sub txtSection_Name_Change()
If bStartCheck Then
    If currentWorkingFile <> txtSection_Name.Text Then
        bSecNameChanged = True
    Else
        bSecNameChanged = False
    End If
End If

syn
End Sub

Public Sub showLimit(bshow As Boolean)
Me.lbLimit.Visible = bshow
End Sub
Public Function openWordDocM(filename As String) As Object ' Document
 Dim retDoc As Object 'Document
 'Dim app As Word.Application
On Error Resume Next

If myWord1 Is Nothing Then
    'If app Is Nothing Then
        Set myWord1 = GetObject(, "Word.Application")
        Debug.Print Err.Number
            
addLog Err.Number
        List1.AddItem Err.Number
        If Err.Number <> 0 Then
            Set myWord1 = CreateObject("Word.Application")
            If myWord1 Is Nothing Then
            
            addLog "Fail to create Word Application"
            End If
            

            Err.Clear
        End If
        End If
        
    'End If
    'myWord1.Visible = True
   
    If InStr(filename, ".dot") <> 0 Or InStr(filename, ".dotx") <> 0 Then
        Set retDoc = myWord1.Documents.Add(filename, True, wdTypeTemplate, True)
       
addLog "Open template"
    Else
        Set retDoc = myWord1.Documents.Add(filename)
        addLog "Add Document called " & filename
        If retDoc Is Nothing Then
           addLog "Warning: The retDoc is null"
        End If
    End If
    
    Dim Spath As String
    Spath = myWord1.Options.DefaultFilePath(8) 'wdStartupPath)
    'app.Run "StartUpMacro"
    
    retDoc.SaveAs filename
    
    addLog "Save Document into " & filename
    
    retDoc.Application.Visible = True
    retDoc.Application.Activate
    retDoc.Application.ShowMe
    Set myAppForTemp = app
    Set openWordDocM = retDoc
    
    'app.ActiveDocument.name = filename
    
End Function


Public Function formatDocxToDoc(filename As String) As String
    On Error Resume Next
    Dim app As Object ' Word.Application '
    Set app = GetObject(, "Word.Application")
    Debug.Print Err.Number
    If Err.Number <> 0 Then
        Set app = CreateObject("Word.Application")
        Err.Clear
    End If
    addLog "Application name=" & app.name
    addLog "Application path=" & app.Path
   ' If InStr(UCase(app.path), "OFFICE12") <> 0 Or InStr(UCase(app.path), "OFFICE13") Then
    If isWord2007OrAbove(app.Path) Then
    addLog "This is Word 2007"
      app.ActiveDocument.SaveFormt = "*.doc"
      Else
      addLog "Return file name =" & filename
      formatDocxToDoc = filename
      Exit Function
    End If
    
    
    addLog "Ready to open " & filename
  
    app.Documents.Open filename

    'Re-show the processed document
    app.ActiveWindow.Visible = False
    filename = replace(filename, ".tmp", ".doc")

    app.ActiveDocument.SaveAs filename, 0
    app.ActiveDocument.Close True
    formatDocxToDoc = filename
  
End Function

