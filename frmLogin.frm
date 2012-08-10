VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4785
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
   ScaleHeight     =   2085
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton btnLogin 
      Caption         =   "&Login"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "&Password"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&Username"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
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


Private Sub btnLogin_Click()
    If (Trim(txtUserName.text) = "") Then
        MsgBox "Please enter the user name"
        txtUserName.SetFocus
    End If
    If (Trim(txtPassword.text) = "") Then
        MsgBox "Please enter the password"
        txtPassword.SetFocus
    End If
    If isFirstLogin = True Then
        addData
    Else
        If getUserData = False Then
            MsgBox "Login Failure. Please try again"
            txtUserName.SetFocus
        Else
            Set ProgramVariable.loggedUser = myUser
            MDIMain.Show
            Unload Me
        End If
        
    End If
End Sub

Private Function getUserData() As Boolean
    Dim temSQL As String
    Dim loginSuccess As Boolean
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblUser"
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                If Trim(txtUserName.text) = mySec.Decode(!UserName, ProgramVariable.SecurityKey) And mySec.Hash(Trim(txtPassword.text)) = !UserPassword Then
                    myUser.UserID = !UserID
                    getUserData = True
                End If
                .MoveNext
            Wend
            .Close
        Else
            addData
        End If
    End With
End Function

Private Sub addData()
    Dim myPerson As New clsPerson
    Dim myRole As New clsRole
    Dim myPrio As clsLeaveType
    
    With myPerson
        .AddedDate = Date
        .AddedTime = Time
        .PersonName = Trim(txtUserName.text)
        .saveData
    End With
    
    With myRole
        .AddedDate = Date
        .AddedTime = Time
        .RoleName = "Administrator"
        .saveData
    End With
    
    With myUser
        .AddedDate = Date
        .AddedTime = Time
        .PersonID = myPerson.PersonID
        .RoleID = myRole.RoleID
        .UserName = mySec.Encode(Trim(txtUserName.text), ProgramVariable.SecurityKey)
        .UserPassword = mySec.Hash(Trim(txtPassword.text))
        .saveData
        .AddedUserID = .UserID
        .saveData
    End With
    
    With myPerson
        .AddedUserID = myUser.UserID
        .saveData
    End With
    
    With myRole
        .AddedUserID = myUser.UserID
        .saveData
    End With
    
    Dim i As Integer
    For i = 1 To 3
        Set myPrio = New clsLeaveType
        
        With myPrio
            .AddedDate = Date
            .AddedTime = Time
            .AddedUserID = myUser.UserID
            Select Case i
                Case 1:            .LeaveTypeName = "Casual"
                Case 2:            .LeaveTypeName = "Sick"
                Case 3:            .LeaveTypeName = "Vacation"
            End Select
            .saveData
        End With
        Set myPrio = Nothing
    Next
    
    Set ProgramVariable.loggedUser = myUser
    MDIMain.Show
    Unload Me
    
End Sub



Private Sub Form_Load()
    Call checkFirstLogin
    SetColours Me
    'GetCommonSettings Me
    Call prepareResize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'SaveCommonSettings Me
End Sub

Private Sub Form_Resize()
  Call frmResize.FormResized(Me)
End Sub

Private Sub prepareResize()
  frmResize.KeepRatio = False
  frmResize.FontResize = True
  Call frmResize.InitializeResizer(Me)
End Sub

Private Sub checkFirstLogin()
    Dim temSQL As String
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblUser"
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
        
        If .RecordCount <= 0 Then
            isFirstLogin = True
        Else
            isFirstLogin = False
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub addRole()

End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Trim(txtUserName.text) <> "" Then
            btnLogin_Click
        Else
            txtUserName.SetFocus
        End If
    End If
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtPassword.SetFocus
    Else
        
    End If
End Sub
