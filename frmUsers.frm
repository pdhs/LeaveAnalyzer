VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmUsers 
   Caption         =   "Users"
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
   Begin VB.CheckBox chkPC 
      Caption         =   "Check1"
      Height          =   495
      Left            =   5760
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtPasswordC 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   14
      Text            =   "111111111111111111111"
      Top             =   2400
      Width           =   2655
   End
   Begin MSDataListLib.DataCombo cmbPerson 
      Height          =   360
      Left            =   5880
      TabIndex        =   6
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   12
      Text            =   "111111111111111111111"
      Top             =   1920
      Width           =   2655
   End
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
      TabIndex        =   17
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   7560
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   6240
      TabIndex        =   15
      Top             =   3120
      Width           =   1215
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
   Begin MSDataListLib.DataCombo cmbRole 
      Height          =   360
      Left            =   6600
      TabIndex        =   10
      Top             =   1440
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "&Role"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "&Confirm Password"
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "&Password"
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "&User Name"
      Height          =   255
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
Attribute VB_Name = "frmUsers"
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
    Dim currentPerson As New clsPerson
    
    Dim rsUsers As New ADODB.Recordset
    Dim rsPersons As New ADODB.Recordset
    
    Dim temSQL As String

Private Sub btnAdd_Click()
    Dim temStr As String
    cmbName.text = Empty
    cmbPerson.text = temStr
    prepareEdit editControls, selectControls
    chkPC.Value = 1
    cmbPerson.SetFocus
End Sub

Private Sub btnCancel_Click()
    clearValues clearControls
    txtPassword.text = "111111111111111111111"
    txtPasswordC.text = "111111111111111111111"
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
    With rsUsers
        If .State = 1 Then .Close
        temSQL = "SELECT tblPerson.PersonName, tblUser.UserID " & _
                    "FROM tblPerson INNER JOIN tblUser ON tblPerson.PersonID = tblUser.PersonID " & _
                    "Where (((tblUser.Deleted) = False) And ((tblPerson.Deleted) = False)) " & _
                    "ORDER BY tblPerson.PersonName"
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
    End With
    With cmbName
        Set .RowSource = rsUsers
        .ListField = "PersonName"
        .BoundColumn = "UserID"
    End With
    Dim Role As New clsFillCombo
    Role.FillSpecificField cmbRole, "Role", "RoleName", True
End Sub

Private Sub fillPersons()
    With rsPersons
        If .State = 1 Then .Close
        temSQL = "SELECT tblPerson.PersonName, tblPerson.PersonID " & _
                    "From tblPerson " & _
                    "Where (((tblPerson.Deleted) = False)) " & _
                    "ORDER BY tblPerson.PersonName "
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
    End With
    With cmbPerson
        Set .RowSource = rsPersons
        .ListField = "PersonName"
        .BoundColumn = "PersonID"
    End With
End Sub

Private Sub btnEdit_Click()
    prepareEdit editControls, selectControls
    cmbRole.SetFocus
    chkPC.Value = 0
    txtUserName.Enabled = False
    cmbPerson.Enabled = False
End Sub

Private Sub btnSave_Click()
    Dim temPerson As New clsPerson
    If IsNumeric(cmbPerson.BoundText) = False Then
        With temPerson
            .AddedDate = Date
            .AddedTime = Time
            .AddedUserID = ProgramVariable.loggedUser.UserID
            .PersonName = cmbPerson.text
            .saveData
            fillPersons
        End With
    Else
        temPerson.PersonID = Val(cmbPerson.BoundText)
    End If
    
    If chkPC.Value = 1 Then
        If txtPassword.text <> txtPasswordC.text Then
            MsgBox "Password and Confirm Password Does not Match"
            txtPassword.SetFocus
            SendKeys "{home}+{end}"
            Exit Sub
        End If
    End If
    
    Dim i As Long
    With current
        If .UserID = 0 Then
            .AddedDate = Date
            .AddedTime = Time
            .AddedUserID = ProgramVariable.loggedUser.UserID
            .UserName = mySec.Encode(Trim(txtUserName.text), ProgramVariable.SecurityKey)
            .PersonID = temPerson.PersonID
        End If
        .RoleID = cmbRole.BoundText
        .UserPassword = mySec.Hash(Trim(txtPassword.text))
        
        .saveData
        
        i = .UserID
        fillNameCombo
        cmbName.BoundText = i
    End With
    prepareSelect editControls, selectControls
End Sub

Private Sub cmbName_Change()
    current.UserID = Val(cmbName.BoundText)
    Call displayDetails
End Sub

Private Sub Form_Load()
    Call setControls
    SetColours Me
    GetCommonSettings Me
    Call prepareResize
    Call fillNameCombo
    Call fillPersons
    prepareSelect editControls, selectControls
End Sub

Private Sub displayDetails()
    clearValues clearControls
    With current
        cmbPerson.BoundText = .PersonID
        cmbRole.BoundText = .RoleID
        txtUserName.text = mySec.Decode(.UserName, ProgramVariable.SecurityKey)
        txtPassword.text = "111111111111111111111"
        txtPasswordC.text = "111111111111111111111"
    End With
End Sub

Private Sub setControls()
    lblName.Caption = "Priorities"
    lblEditName.Caption = "LeaveType"
    Me.Caption = "Manage Priorities"
    
    With editControls
        .Add cmbPerson
        .Add cmbRole
        .Add txtUserName
        .Add txtPassword
        .Add txtPasswordC
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
        .Add cmbPerson
        .Add cmbRole
        .Add txtUserName
        .Add txtPassword
        .Add txtPasswordC
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



Private Sub txtPassword_Change()
    chkPC.Value = 1
End Sub

Private Sub txtPassword_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtPasswordC_Change()
    chkPC.Value = 1
End Sub

Private Sub txtPasswordC_GotFocus()
    SendKeys "{home}+{end}"
End Sub
