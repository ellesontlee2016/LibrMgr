VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTemplate 
   Caption         =   "Template Manager"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   5160
   HelpContextID   =   1001
   Icon            =   "frmTemplate.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   5160
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   600
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   550
      Width           =   852
   End
   Begin VB.TextBox txtAddName 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3852
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "Edit Doc"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   550
      Width           =   1092
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   550
      Width           =   852
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   550
      Width           =   852
   End
   Begin VB.ComboBox cmbTemplateList 
      Height          =   315
      ItemData        =   "frmTemplate.frx":1CCA
      Left            =   1200
      List            =   "frmTemplate.frx":1CCC
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3852
   End
   Begin VB.Label Label1 
      Caption         =   "Template List"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1092
   End
End
Attribute VB_Name = "frmTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim buff() As Byte
Dim bLoadNewDoc As Boolean
Dim bSystemScope As Boolean
'Dim myApp As Word.Application
Dim currentWorkingDoc As String
Dim currentDocName As String
Dim currentDoc As Object ' Word.Document
Dim currentOpenDotDate As Date
Dim bOpenDoc As Boolean
Dim bDocChanged As Boolean
Dim bForcedByMDIP As Boolean
Dim startStamp As Date
Dim closeStamp As Date
Dim bCheckTime As Boolean

'Public WithEvents myApp As Word.Application
Public myApp As Object




Private Function isDuplicateKey(Key As String) As Boolean
    Dim no, i As Integer
    no = Me.cmbTemplateList.ListCount
    
    For i = 0 To no - 1
        If CStr(Me.cmbTemplateList.List(i)) = Key Then
          isDuplicateKey = True
          Exit Function
        End If
    Next i
    isDuplicateKey = False
End Function

Private Sub cmdAdd_Click()
 Dim docName As String
 Dim startStrg As String
 Dim secName As String
    Dim c As Integer
   Dim y As Integer
   Dim sFile As String
   Dim tmp As String
    Dim i As Integer

    Dim FileArray() As String

 
    If Me.cmdAdd.Caption = "Add" Then
        bLoadNewDoc = False
        Me.cmbTemplateList.Visible = False
        Me.cmdDelete.Visible = False
        Me.cmdView.Visible = False
        Me.cmdCancel.Visible = True
        Me.cmbTemplateList.Visible = False
        Me.txtAddName.Visible = True
        Me.txtAddName.Enabled = True
        Me.txtAddName.Text = "<New Template Name>"
        
        cmdAdd.Caption = "OK"
        CommonDialog1.MaxFileSize = 4096
        CommonDialog1.DefaultExt = "dot"
        If getXMLOfficeFiles() Then
                 CommonDialog1.Filter = "Word Template (*.dot;*.dotx)|*.dot;*.dotx"
                 CommonDialog1.filename = "*.dot;*.dotx"
                 CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
                 CommonDialog1.CancelError = True
                 On Error GoTo handleError
                
                 CommonDialog1.ShowOpen  ' = 1
                 
           
                 If Right(CommonDialog1.filename, 5) = "*.dot" Or Right(CommonDialog1.filename, 6) = "*.dotx" Then
        
                    Screen.MousePointer = vbDefault
                    Me.cmbTemplateList.Visible = True
                    Me.cmdDelete.Visible = True
                    Me.cmdView.Visible = True
                    Me.txtAddName.Visible = False
                    Me.cmdCancel.Visible = False
                    Me.cmdAdd.Caption = "Add"
                    loadInfo
                    bSystemScope = True
                 Exit Sub
                 End If
                 
                 
                  If Right(CommonDialog1.filename, 4) <> ".dot" And Right(CommonDialog1.filename, 5) <> ".dotx" Then
                   MsgBox "ERROR: Illegal file type; only DOT/DOTX template files can be imported into the document library.", vbOKOnly, "Error"
                   Screen.MousePointer = vbDefault
                    Me.cmbTemplateList.Visible = True
                    Me.cmdDelete.Visible = True
                    Me.cmdView.Visible = True
                    Me.txtAddName.Visible = False
                    Me.cmdCancel.Visible = False
                    Me.cmdAdd.Caption = "Add"
                    loadInfo
                    bSystemScope = True
                 Exit Sub
                 End If
                 
                 
                 
            Else
             CommonDialog1.Filter = "Word Template (*.dot)|*.dot"
             CommonDialog1.filename = "*.dot"
             CommonDialog1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
             CommonDialog1.CancelError = True
             On Error GoTo handleError
             CommonDialog1.ShowOpen     ' = 1
              If Right(CommonDialog1.filename, 5) = "*.dot" Then
        
                Screen.MousePointer = vbDefault
                Me.cmbTemplateList.Visible = True
                Me.cmdDelete.Visible = True
                Me.cmdView.Visible = True
                Me.txtAddName.Visible = False
                Me.cmdCancel.Visible = False
                Me.cmdAdd.Caption = "Add"
                loadInfo
                bSystemScope = True
                Exit Sub

             End If
             
                         
             
             If Right(CommonDialog1.filename, 4) <> ".dot" Then
                Screen.MousePointer = vbDefault
                MsgBox "ERROR: Illegal file type; only Dot Template files can be imported into the document library.", vbOKOnly, "Error"
                
                Screen.MousePointer = vbDefault
                Me.cmbTemplateList.Visible = True
                Me.cmdDelete.Visible = True
                Me.cmdView.Visible = True
                Me.txtAddName.Visible = False
                Me.cmdCancel.Visible = False
                Me.cmdAdd.Caption = "Add"
                loadInfo
                bSystemScope = True
                Exit Sub
             End If
            
             
             
           End If
        
        startStrg = CommonDialog1.filename & Chr(0) & Chr(0)
        
        For c = 1 To Len(CommonDialog1.filename)
        
            sFile = StripItem(startStrg, Chr(0))
            If sFile = "" Then Exit For
               ReDim Preserve FileArray(0 To c - 1)
               FileArray(y) = LCase(sFile)
               y = y + 1
            Next
        
        Dim prefix As String
        
        
        
        
        If c = 1 Then
        restore
        Exit Sub
        End If
    If c = 2 Then
        docName = FileArray(0)
        If docName <> "" Then
            secName = Mid(docName, InStrRev(docName, "\") + 1, InStrRev(docName, ".") - InStrRev(docName, "\") - 1)
            If Not isDuplicateKey(secName) Then

                buff = OpenToDBNewTemplate(FileArray(0))
                'bufferCollect.Add buffer
                saveToTempTbl secName, buff
                bLoadNewDoc = True
                bSystemScope = False
            End If
                
        End If
    Else
    
    prefix = FileArray(0)
    For i = 1 To c - 2
    incD = CDbl(i)
        docName = FileArray(i)
        If docName <> "" Then
            secName = Mid(docName, InStrRev(docName, "\") + 1, InStrRev(docName, ".") - InStrRev(docName, "\") - 1)
            If Not isDuplicateKey(secName) Then
                buff = OpenToDBNewTemplate(prefix & "\" & docName)
                saveToTempTbl secName, buff
                'bufferCollect.Add buffer
                bLoadNewDoc = True
                bSystemScope = False
            End If
                
        End If
    Next i
    End If
    
  '      docName = CommonDialog1.FileName
  '      If docName <> "" Then
  '      buff = openToDB(docName)
  '      Me.txtAddName.Text = Mid(docName, InStrRev(docName, "\") + 1, InStrRev(docName, ".") - InStrRev(docName, "\") - 1)
 '       bLoadNewDoc = True
' '       'ReDim buff(0)
 '       bSystemScope = False
        
  '      Else
  '       restore
  '      End If
        
        Me.cmbTemplateList.Visible = True
        Me.cmdDelete.Visible = True
        Me.cmdView.Visible = True
        Me.txtAddName.Visible = False
        Me.cmdCancel.Visible = False
        Me.cmdAdd.Caption = "Add"
        loadInfo
        bSystemScope = True
    Else
    
        If isDuplicateKey(Me.txtAddName.Text) Or Me.txtAddName.Text = "" Then
         MsgBox "Duplicated Template Name or empty Template Name!", , "Warning"
         Me.txtAddName.SetFocus
         Exit Sub
        End If
        
        If Not bLoadNewDoc Then
            MsgBox "Please select a document", , "Warning"
            Exit Sub
        End If
        
        'If DataEnvironment2.cnnSection.State <> adStateOpen Then
        'DataEnvironment2.cnnSection.Open
        'End If
        
        On Error Resume Next
        
        saveToTempTbl Me.txtAddName, buff
        'DataEnvironment2.cmdInsertTemp Me.txtAddName, buff
        
       'If DataEnvironment2.cnnSection.State = adStateOpen Then
        '    DataEnvironment2.cnnSection.Close
        'End If
    
        Me.cmbTemplateList.Visible = True
        Me.cmdDelete.Visible = True
        Me.cmdView.Visible = True
        Me.txtAddName.Visible = False
        Me.cmdCancel.Visible = False
        Me.cmdAdd.Caption = "Add"
        loadInfo
        bSystemScope = True
    
        
    End If
    Exit Sub
    
handleError:

    Screen.MousePointer = vbDefault
    Me.cmbTemplateList.Visible = True
    Me.cmdDelete.Visible = True
    Me.cmdView.Visible = True
    Me.txtAddName.Visible = False
    Me.cmdCancel.Visible = False
    Me.cmdAdd.Caption = "Add"
    loadInfo
    bSystemScope = True
    Exit Sub
    
    End Sub
    
    


Public Sub saveToTempTbl(tempName As String, buff() As Byte)
    If DataEnvironment2.cnnSection.State <> adStateOpen Then
    getpwdConn DataEnvironment2.cnnSection
    DataEnvironment2.cnnSection.Open
    End If
    
    On Error Resume Next
    
    
    DataEnvironment2.cmdInsertTemp tempName, buff
    
    If DataEnvironment2.cnnSection.State = adStateOpen Then
        DataEnvironment2.cnnSection.Close
    End If
   
End Sub

Private Sub restore()
    Me.cmbTemplateList.Visible = True
    Me.txtAddName.Visible = False
    'Me.lbdoc.Visible = False
    Me.cmdAdd.Visible = True
    Me.cmdCancel.Visible = False
    Me.cmdAdd.Caption = "Add"
    Me.cmdView.Caption = "Edit Doc"
    Me.cmdDelete.Visible = Me.cmbTemplateList.ListCount <> 0
    loadInfo
    Me.cmdView.Visible = cmbTemplateList.ListCount <> 0
    Me.cmdView.Enabled = cmbTemplateList.ListCount <> 0
    Me.cmdDelete.Enabled = cmbTemplateList.ListCount <> 0
End Sub

Private Sub cmdCancel_Click()
    bDocChanged = False
    bOpenDoc = False
    
    
    
    
    If Not bSystemScope Then
        restore
        bSystemScope = True
        Me.cmdCancel.Visible = False
    Else
        Me.Hide
        
    End If
    
    On Error GoTo handle
    myApp.VisualBasic
    Exit Sub
handle:
Set myApp = Nothing

End Sub


Private Sub cmdView_Click()
    
    If Me.cmdView.Caption = "Edit Doc" Then
        Me.cmbTemplateList.Visible = False
        Me.cmdDelete.Visible = False
        Me.cmdAdd.Visible = False
        Me.cmdCancel.Visible = True
        Me.cmbTemplateList.Visible = False
        Me.txtAddName.Visible = True
        cmdView.Caption = "OK"
        cmdView.Enabled = False
        Me.txtAddName.Text = Me.cmbTemplateList.Text
        Me.txtAddName.Enabled = False
        bSystemScope = False
        viewDoc
        bOpenDoc = True
         Me.Timer1.Enabled = True
         bCheckTime = False

  
    Else
        bOpenDoc = False
        bDocChanged = False
        If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
        End If
        
        If bLoadNewDoc Then
              DataEnvironment2.cmdUpdateTemp buff, getClientName(), Now(), Me.txtAddName.Text
        Else
        DataEnvironment2.cmdUpdateTemp OpenToDBNewTemplate(currentWorkingDoc), getClientName(), Now(), Me.txtAddName.Text
        End If
        
        If DataEnvironment2.cnnSection.State = adStateOpen Then
            DataEnvironment2.cnnSection.Close
        End If

        On Error Resume Next
        myApp.ActiveDocument.Close
        Me.cmbTemplateList.Visible = True
        Me.cmdDelete.Visible = True
        Me.cmdAdd.Visible = True
        Me.txtAddName.Visible = False
        'Me.lbdoc.Visible = False
        Me.cmdView.Caption = "Edit Doc"
        Me.cmdCancel.Visible = False
        'Me.OLE1.Visible = True
        'Me.lbClick.Visible = True
        'Me.lbClick.Caption = "Click blue icon to view dot"
        Me.txtAddName.Enabled = True
        'Me.lbdoc.Visible = True
        'Me.cmdBrowse.Visible = True
        bSystemScope = True
        Dim nowDate As Date
        'If Not currentDoc Is Nothing Then
        
        On Error GoTo NextI
        
        GetFileTimeInfo currentDoc.fullname, , nowDate
        GetFileTimeInfo currentDoc.fullname, , startStamp

        If nowDate <> currentOpenDotDate Then
        'currentDoc.AcceptAllRevisions
         buff = OpenToDBNewTemplate(currentDoc.fullname)
               
               
        End If

        GoTo NextII
        
NextI:

        GetFileTimeInfo currentDocName, , nowDate
        If nowDate <> currentOpenDotDate Then
        'currentDoc.AcceptAllRevisions
         buff = OpenToDBNewTemplate(currentDocName)
        End If
NextII:
    
        If nowDate <> currentOpenDotDate Then
            saveToTempTbl Me.txtAddName, buff
            Set currentDoc = Nothing
            
        End If
        
        loadInfo
        
    End If


End Sub

Private Sub loadInfo()
    Dim def As String
    Dim counter As Integer
    Me.cmbTemplateList.Clear
    
    If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
    End If
    DataEnvironment2.cmdGetTempInfo
    counter = 0
    def = ""
    While Not DataEnvironment2.rscmdGetTempInfo.EOF
        If counter = 0 Then
            def = DataEnvironment2.rscmdGetTempInfo!Template_Name
        End If
        If Not def = "" Then
            Me.cmbTemplateList.AddItem DataEnvironment2.rscmdGetTempInfo!Template_Name
            counter = counter + 1
            Else
            DataEnvironment2.cmdDeleteTemp ""
        End If
        DataEnvironment2.rscmdGetTempInfo.MoveNext
    Wend
    
     If Me.cmbTemplateList.ListCount > 0 Then
        Me.cmbTemplateList.Text = def
     End If
     'Me.lbClick.Visible = counter <> 0
     'Me.OLE1.Visible = counter <> 0
     Me.cmdView.Enabled = counter <> 0
     Me.cmdDelete.Enabled = counter <> 0
    
    If DataEnvironment2.cnnSection.State = adStateOpen Then
        DataEnvironment2.cnnSection.Close
    End If
    

End Sub

Private Sub cmdDelete_Click()
    If Not Me.cmbTemplateList.Text = "" Then
        If Not MsgBox("Are you sure you want to delete this template?", vbYesNo, "Delete Template") = vbYes Then
            Exit Sub
        End If
        If DataEnvironment2.cnnSection.State <> adStateOpen Then
            getpwdConn DataEnvironment2.cnnSection
            DataEnvironment2.cnnSection.Open
        End If
        
        DataEnvironment2.cmdDeleteTemp Me.cmbTemplateList.Text
        If DataEnvironment2.cnnSection.State = adStateOpen Then
            DataEnvironment2.cnnSection.Close
        End If
        restore
    End If
End Sub



Private Sub Form_Activate()
bTempActive = True
End Sub

Private Sub Form_Deactivate()
bTempActive = False
If Not bSectionUnload Then
bSectionActive = True
End If


End Sub

Private Sub Form_GotFocus()
bTempActive = True
bSectionActive = False

End Sub

Private Sub Form_Load()
'Me.HelpContextID = ID_TEMPLATEMGR
    loadInfo
    'Me.lbdoc.Visible = False
    Me.cmdCancel.Visible = False
    Set myApp = Nothing
    bSystemScope = True
    bOpenDoc = False
    bTempUnload = False
    bForcedByMDIP = False
End Sub



Private Sub mnuHelp_Click()

End Sub


Private Sub Form_LostFocus()
bTempActive = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

 If bForcedByMDIP Then
 Exit Sub
 End If
 

 If MsgBox("Do you want to close Template Manager?", vbYesNo, "Confirmation") = vbYes Then
   
   Else
   Cancel = 1
 End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
bTempActive = False
bTempUnload = True
If Not bSectionUnload Then
  bSectionActive = True
End If

End Sub

Public Sub shutdownByParent()
bForcedByMDIP = True
End Sub

Private Sub myApp_DocumentBeforeSave(ByVal Doc As Object, SaveAsUI As Boolean, Cancel As Boolean) ' Word.Document, SaveAsUI As Boolean, Cancel As Boolean)
Exit Sub
 'If Not currentDoc Is Nothing And bOpenDoc Then
'''' If bOpenDoc Then
''''  Dim nowDate As Date
  
  
   If bOpenDoc Then ' And myWord1.ActiveDocument.Saved Then
       'Doc.AcceptAllRevisionsShown
       If SaveAsUI Then
       'Doc.SaveAs (Doc.FullName)
       Else
       Doc.Save
       End If
         currentDocName = Doc.fullname
  Set currentDoc = Doc

    End If
  
  
 '''' doc.SaveAs (doc.FullName)
  'bDocChanged = True
  'cmdView.Enabled = bDocChanged
  ''''bDocChanged = True
''''  syn
 '''' End If
  
End Sub

Private Sub myApp_DocumentChange()
'MsgBox "Hey you"
On Error Resume Next
currentDocName = myApp.ActiveDocument.fullname
End Sub

Private Sub OLE1_Click()
    viewDoc
End Sub

Private Sub syn()
    Me.cmdView.Enabled = bDocChanged
End Sub

Public Function getFileWithDefaultPath(filename As String) As String
On Error Resume Next
    Set myApp = GetObject(, "Word.Application")
    Debug.Print Err.Number
    If Err.Number <> 0 Then
        Set myApp = CreateObject("Word.Application")
        
        Err.Clear
    End If
    myApp.Visible = True
    defaultDocPath = myApp.Options.DefaultFilePath(8) 'wdDocumentsPath)
    If InStrRev(filename, "\") <> 0 Then
        filename = defaultDocPath & "\" & Right(filename, Len(filename) - InStrRev(filename, "\"))
    Else
        filename = defaultDocPath & "\" & filename
End If
    getFileWithDefaultPath = filename
End Function


Private Sub viewDoc()
    Dim fh As Integer
    
    If DataEnvironment2.cnnSection.State <> adStateOpen Then
        getpwdConn DataEnvironment2.cnnSection
        DataEnvironment2.cnnSection.Open
    End If
    
    If Me.cmbTemplateList.Text <> "" Then
        DataEnvironment2.cmdGetTempDoc Me.cmbTemplateList.Text
    Else
        Exit Sub
    End If
    
    If DataEnvironment2.rscmdGetTempDoc.RecordCount > 0 Then
        
        buff = DataEnvironment2.rscmdGetTempDoc!Word_Doc
        fh = FreeFile()
        
        If DataEnvironment2.cnnSection.State = adStateOpen Then
            DataEnvironment2.cnnSection.Close
        End If
    
        currentWorkingDoc = CreateTempFile()
        currentWorkingDoc = Left(currentWorkingDoc, InStrRev(currentWorkingDoc, ".")) & "dot"
        Open currentWorkingDoc For Binary As fh
        Put fh, , buff
        
        
        Close fh
        GetFileTimeInfo currentWorkingDoc, , currentOpenDotDate
        
        'getFileWithDefaultPath ""
        Set currentDoc = openWordDoc(myApp, currentWorkingDoc)
        
        
       ' On Error Resume Next
        Set myApp = getApp()
        'openWordDoc tmpstr
        
    End If
    
   
End Sub
Private Sub Timer1_Timer()
Dim tmpDate As Date
   If Not bCheckTime Then
      GetFileTimeInfo currentWorkingDoc, , startStamp
    bCheckTime = True
    End If
    
    
    If GetFileTimeInfo(currentWorkingDoc, tmpDate, closeStamp) Then
      ''''  Me.lbreateTime.Caption = CStr(startStamp)
      ''''  Me.lbWriteTime.Caption = CStr(closeStamp)
    
        If bOpenDoc Then
                 bDocChanged = DateDiff("s", startStamp, closeStamp)
                
                If bDocChanged Then
                    bOpenDoc = False
                    Me.cmdView.Enabled = True
                Else
                   Me.cmdView.Enabled = False
                End If
           ' End If
        
        End If
    End If
End Sub

Private Sub myApp_Quit()
Set myApp = Nothing
End Sub
