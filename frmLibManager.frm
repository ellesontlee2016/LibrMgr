VERSION 5.00
Begin VB.MDIForm frmLibManager 
   BackColor       =   &H8000000C&
   Caption         =   "Library Manager"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10830
   Icon            =   "frmLibManager.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu manuItem_File_Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu manuTables 
      Caption         =   "&Tables"
      Begin VB.Menu menuTables_Template 
         Caption         =   "Tem&plate"
      End
      Begin VB.Menu menuItemTables_Section 
         Caption         =   "&Section"
      End
   End
   Begin VB.Menu menuWindow 
      Caption         =   "&Window"
      Begin VB.Menu menuWindow_HZ 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu menuWindow_VT 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu menuWindow_Cascade 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmLibManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim frmSection As frmStoreDoc
Dim frmSection As frmStoreDocNew
Dim frmTemp As frmTemplate
Dim Directory_Root As String
Dim userkey As String
Dim checker As New ExpirationChecker
'Dim uniqueKey As String

Private Sub manuItem_File_Exit_Click()

 If MsgBox("Do you want to leave Library Manager?", vbYesNo, "Library Manager Exit Confirmation") = vbYes Then
   End
 End If
 
End Sub



Private Sub getInputParameter()
    Dim sArgs() As String
    SHOW_DEBUG = 0
    sArgs = Split(Command$, " ") 'Command holds your command line arguments
    For i = 0 To UBound(sArgs)
        Select Case LCase(sArgs(i))
            Case "-D", "-d"
            SHOW_DEBUG = 1
            Case "-KEY", "-key"
            userkey = sArgs(i + 1)
         Case Else
        End Select
     Next
    
End Sub


Private Sub MDIForm_Load()
    'uniqueKey = mySysInfo.getUniqueTracingHashKey()

    getInputParameter
    
    Directory_Root = GetRegValue(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\CorsPro\Settings", "Directory_Company")
    ''app.HelpFile = Directory_Root & "PQManager\LibrMgr.chm"
    app.HelpFile = app.path & "\LibrMgr.chm"
    
    bTempUnload = False
    bSectionUnload = False

    bconvertEn = False
    bconvertAdo = False
    
   ' Set frmSection = New frmStoreDoc
    Set frmSection = New frmStoreDocNew
    frmSection.setParent Me
  Set frmTemp = New frmTemplate
  
  LibModule.choseDB
      frmTemp.Show

    frmSection.setConnStr
    frmSection.Show
    frmSection.Refresh
    bTempActive = False
    bSectionActive = True
    

'    Dim constr As String
'    constr = DataEnvironment2.cnnSection.ConnectionString
 '    If InStr(1, constr, "\ProposalDB.mdb", vbTextCompare) = 0 Then
'        If Not bconvertEn Then
        
 '        Dim value As String
 '        value = GetRegValue(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\CorsPro\Settings", "Directory_Company")
 '           DataEnvironment2.cnnSection.ConnectionString = replaceDataBase(DataEnvironment2.cnnSection.ConnectionString, "Data Source=" & value & "PropGen\Data\ProposalDB.mdb")
 '           bconvertEn = True
 '       End If
                
    '            DataEnvironment2.cnnSection.ConnectionString = DataEnvironment2.cnnSection.ConnectionString & _
    '                    ";Data Source=" & app.path & "\Data\ProposalDB.mdb"
'    frmTemp.Show

'    frmSection.setConnStr
'    frmSection.Show
 '   frmSection.Refresh
'    End If

   ''  If myChecker.isTrailVersion() Then
    ' Me.Caption = Me.Caption & " (Trial Version)"
   ' End If
    Dim PopMsg As String
   ' If myChecker.isExpiredDB(uniqueKey) Then
    If isExpiredNew() Then
     '   If myChecker.isTrailVersion() Then
     '       PopMsg = "Your software trial has expired.  Please contact Cors Productivity Solutions, Inc. at support@corspro.com to purchase a license for this software or for any other assistance."
     '   Else
            PopMsg = "Your license to use this software has expired.  Please contact Cors Productivity Solutions, Inc. at support@corspro.com to renew your license or for any other assistance."
     '   End If
      MsgBox PopMsg
      End
    End If
    
     Me.Arrange vbCascade
     
     
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If MsgBox("Do you want to leave Library Manager?", vbYesNo, "Library Manager Exit Confirmation") = vbYes Then
   If Not frmSection Is Nothing Then
    frmSection.shutdownByParent
   End If
  
      If Not frmTemp Is Nothing Then
      
    frmTemp.shutdownByParent
   End If

   Else
   Cancel = 1
 End If

End Sub

Private Sub MDIForm_Resize()
 'Me.Arrange vbTileHorizontal
End Sub






Private Sub menuHelp_Click()
Dim hwndHelp As Long


If Not bTempActive And Not bSectionActive Then
''hwndHelp = HtmlHelp(hWnd, Directory_Root & "PQManager\LibrMgr.chm", HH_DISPLAY_TOPIC, 0)
hwndHelp = HtmlHelp(hWnd, app.path & "\LibrMgr.chm", HH_DISPLAY_TOPIC, 0)
ElseIf bTempActive Then
''hwndHelp = HtmlHelp(hWnd, Directory_Root & "PQManager\LibrMgr.chm", HH_HELP_CONTEXT, 1001)
hwndHelp = HtmlHelp(hWnd, app.path & "\LibrMgr.chm", HH_HELP_CONTEXT, 1001)

ElseIf bSectionActive Then
''hwndHelp = HtmlHelp(hWnd, Directory_Root & "PQManager\LibrMgr.chm", HH_HELP_CONTEXT, 1000)
hwndHelp = HtmlHelp(hWnd, app.path & "\LibrMgr.chm", HH_HELP_CONTEXT, 1000)


End If



'Me.DisplayHelp 1000
' If  = frmStoreDoc Then
'hwndHelp = HtmlHelp(hWnd, Directory_Root & "PropGen\LibrMgr.chm", HH_HELP_CONTEXT, 1000)
' ElseIf Me.ActiveForm = frmTemplate Then
 'hwndHelp = HtmlHelp(hWnd, Directory_Root & "PropGen\LibrMgr.chm", HH_HELP_CONTEXT, 1001)
' Else
' hwndHelp = HtmlHelp(hWnd, Directory_Root & "PropGen\LibrMgr.chm", HH_DISPLAY_TOPIC, 0)
' End If
 
 
 
 

'DisplayHelp
'    Dim Directory_Root As String
'    Directory_Root = GetRegValue(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\CorsPro\Settings", "Directory_Company")
    
'Dim hwndHelp As Long
'   hwndHelp = HtmlHelp(hWnd, Directory_Root & "PropGen\LibrMgr.chm", HH_DISPLAY_TOPIC, 0)


End Sub

Private Sub menuItemTables_Section_Click()
    If frmSection Is Nothing Then
     'Set frmSection = New frmStoreDoc
     Set frmSection = New frmStoreDocNew
     End If
    frmSection.Show
bTempActive = False
bSectionActive = True

    frmSection.SetFocus
End Sub

Private Sub menuTables_Template_Click()
    If frmTemp Is Nothing Then
     Set frmTemp = New frmTemplate
    
     End If
    frmTemp.Show
    bTempActive = True
bSectionActive = False

   
    frmTemp.SetFocus
End Sub

Private Sub menuWindow_Cascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub menuWindow_HZ_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub menuWindow_VT_Click()
    Me.Arrange vbTileVertical
End Sub

Public Sub DisplayHelp(Optional ContextID As Long = 0)
    Dim Directory_Root As String
    ''Directory_Root = GetRegValue(HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\CorsPro\Settings", "Directory_Company")
    ''app.HelpFile = Directory_Root & "PQManager\LibrMgr.chm"
    app.HelpFile = app.path & "\LibrMgr.chm"
  
  
    'Application.Help chm
    
    'Application.Help Directory_Root & "PropGen\LibrMgr.chm"
    'app.Help
End Sub

