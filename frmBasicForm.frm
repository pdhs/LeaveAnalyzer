VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmBasicForm 
   Caption         =   "Priority"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   9555
   Begin MSDataListLib.DataCombo cmbName 
      Height          =   3540
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6244
      _Version        =   393216
      Style           =   1
      Text            =   ""
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "C&lose"
      Height          =   495
      Left            =   8160
      TabIndex        =   11
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   7560
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtComments 
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblComments 
      Caption         =   "Co&mments"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblEditName 
      Caption         =   "&Name"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblName 
      Caption         =   "&Name"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmBasicForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Auther : Dr. M. H. B. Ariyaratne
'          buddhika.ari@gmail.com
'          buddhika_ari@yahoo.com
'          +94 71 58 12399
'          GPL Licence

Option Explicit
    Dim frmResize As New clsResizer
    Dim FSys As New Scripting.FileSystemObject
    Dim isFirstLogin As Boolean
    
    Dim mySec As New clsSecurity
    Dim myUser As New clsUser
    
    Dim editControls As New Collection
    Dim selectControls As New Collection
    Dim clearControls As New Collection
    
    Dim current As New clsPriority
    

Private Sub btnAdd_Click()
    Dim temStr As String
    cmbName.text = Empty
    txtName.text = temStr
    prepareEdit editControls, selectControls
    txtName.SetFocus
End Sub

Private Sub btnCancel_Click()
    clearValues clearControls
    prepareSelect editControls, selectControls
    cmbName.SetFocus
    displayDetails
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim a As Integer
    a = MsgBox("Are you sure you want to delete " & cmbName.text & "?", vbYesNo)
    If a = vbYes Then
        current.Deleted = True
        current.DeletedDate = Date
        current.DeletedTime = Now
        current.DeletedUserID = ProgramVariable.loggedUser.UserID
        current.saveData
        MsgBox "Deleted"
        fillNameCombo
        cmbName.text = Empty
    End If
End Sub

Private Sub fillNameCombo()
    Dim allItems As New clsFillCombo
    allItems.FillSpecificFieldOrder cmbName, "Priority", "PriorityName", "OrderNo", True
End Sub

Private Sub btnEdit_Click()
    prepareEdit editControls, selectControls
    txtName.SetFocus
End Sub

Private Sub btnSave_Click()
    Dim i As Long
    With current
        .PriorityName = txtName.text
        .OrderNo = Val(txtComments.text)
        .saveData
        i = .PriorityID
        fillNameCombo
        cmbName.BoundText = i
    End With
    prepareSelect editControls, selectControls
End Sub

Private Sub cmbName_Change()
    current.PriorityID = Val(cmbName.BoundText)
    Call displayDetails
End Sub

Private Sub Form_Load()
    Call setControls
    SetColours Me
    GetCommonSettings Me
    Call prepareResize
    Call fillNameCombo
End Sub

Private Sub displayDetails()
    clearValues clearControls
    With current
        txtName.text = .PriorityName
        txtComments.text = .OrderNo
    End With
End Sub

Private Sub setControls()
    lblName.Caption = "Priorities"
    lblEditName.Caption = "Priority"
    Me.Caption = "Manage Priorities"
    
    With editControls
        .Add txtName
        .Add txtComments
        .Add btnSave
        .Add btnCancel
    End With
    
    With selectControls
        .Add cmbName
        .Add btnAdd
        .Add btnEdit
        .Add btnDelete
    End With

    With clearControls
        .Add txtName
        .Add txtComments
    End With
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub

Private Sub Form_Resize()
  Call frmResize.FormResized(Me)
End Sub

Private Sub prepareResize()
  frmResize.KeepRatio = False
  frmResize.FontResize = True
  Call frmResize.InitializeResizer(Me)
End Sub



