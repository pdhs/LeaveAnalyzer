VERSION 5.00
Begin VB.Form frmMyDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Priority"
   ClientHeight    =   3090
   ClientLeft      =   3120
   ClientTop       =   4200
   ClientWidth     =   7935
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7935
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox txtCPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "C&lose"
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtComments 
      Height          =   375
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "&Password"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ConfirmPassword"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblComments 
      Caption         =   "User Name"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblEditName 
      Caption         =   "&Name"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMyDetails"
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
    
    Dim current As New clsUser
    
    Dim passwordChanged As Boolean
    
    
Private Sub btnClose_Click()
    Unload Me
End Sub


Private Sub btnSave_Click()
    Dim i As Long
    Dim currentPerson As New clsPerson
    If txtPassword.text <> txtCPassword.text Then
        MsgBox "Confirm Password not Matching"
        txtCPassword.SetFocus
        Exit Sub
    End If
    If Trim(txtName.text) = "" Then
        MsgBox "Name ?"
        txtName.SetFocus
        Exit Sub
    End If
    current.UserID = ProgramVariable.loggedUser.UserID
    currentPerson.PersonID = current.PersonID
    With currentPerson
        .PersonName = txtName.text
        .saveData
    End With
    
    If Trim(txtPassword.text) <> "" Then
        With current
            .UserPassword = mySec.Hash(txtPassword.text)
            .saveData
        End With
    Else

    End If
    
    MsgBox "Saved"
    
End Sub

Private Sub Form_Load()
    Call setControls
    SetColours Me
    GetCommonSettings Me
    Call prepareResize
    Call displayDetails
End Sub

Private Sub displayDetails()
    Dim currentPerson As New clsPerson
    
    With current
        .UserID = ProgramVariable.loggedUser.UserID
        txtComments.text = mySec.Decode(.UserName, ProgramVariable.SecurityKey)
    End With
    
    With currentPerson
        .PersonID = current.PersonID
        txtName.text = .PersonName
    End With
End Sub

Private Sub setControls()
    lblEditName.Caption = "Name"
    Me.Caption = "Manage My Details"
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


