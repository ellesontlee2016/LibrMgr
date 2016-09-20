VERSION 5.00
Begin VB.Form frmConfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4215
   Icon            =   "frmConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Doc Type"
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   3975
      Begin VB.ComboBox cmbDocType 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Location of inserted section"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3975
      Begin VB.OptionButton optAfter 
         Caption         =   "AFTER the selected row"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2775
      End
      Begin VB.OptionButton OptBefore 
         Caption         =   "BEFORE the selected row"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Library Section's Object Type"
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton optExternal 
         Caption         =   "Excel range (EXTERNAL to doc library)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   3735
      End
      Begin VB.OptionButton optInternal 
         Caption         =   "Word document (INTERNAL to doc library)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bOK As Boolean
Dim bIsInternal As Boolean
Dim bIsBefore As Boolean
Dim myDocType As String
Private Sub cmdCancel_Click()
bOK = False
Me.OptBefore.value = True
Me.optInternal.value = True
 Me.Hide
End Sub

Private Sub cmdOK_Click()
bOK = True
myDocType = Me.cmbDocType.Text
Me.Hide
End Sub

Public Function getDocType() As String
 getDocType = myDocType
End Function
Public Function isOKClick() As Boolean
isOKClick = bOK
End Function

Public Function isInternal() As Boolean
    isInternal = bIsInternal
End Function

Public Function isInsertBefore() As Boolean
    isInsertBefore = bIsBefore
End Function

Private Sub Form_Load()
Dim mycol As Collection
bIsInternal = True
bIsBefore = True
bOK = False
Set mycol = populateDocType()
For i = 1 To mycol.Count
Me.cmbDocType.AddItem mycol.item(i)
Next
Me.cmbDocType.Text = Me.cmbDocType.List(0)

End Sub

Private Sub optAfter_Click()
bIsBefore = Not optAfter.value
End Sub

Private Sub OptBefore_Click()
bIsBefore = Not optAfter.value
End Sub

Private Sub optExternal_Click()
 bIsInternal = Not optExternal.value
End Sub

Private Sub optInternal_Click()
bIsInternal = Not optExternal.value
End Sub


