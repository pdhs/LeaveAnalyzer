VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmLeave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Letter"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9630
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
   ScaleHeight     =   4665
   ScaleWidth      =   9630
   Begin VB.CheckBox chkHalfDay 
      Caption         =   "Half-Day"
      Height          =   255
      Left            =   3600
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   8760
      Top             =   840
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "C&lose"
      Height          =   495
      Left            =   8160
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtLeaveID 
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkPlanned 
      Caption         =   "Planned"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtComment 
      Height          =   1935
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1920
      Width           =   7575
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpLeaveDate 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   118095875
      CurrentDate     =   40940
   End
   Begin MSDataListLib.DataCombo cmbEmployee 
      Height          =   360
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbLeaveType 
      Height          =   360
      Left            =   1800
      TabIndex        =   12
      Top             =   1440
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label lblNewFrom 
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "Employee"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Leave Data"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Type"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Comment"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
End
Attribute VB_Name = "frmLeave"
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
    Dim myLeave As New clsLeave
    Dim myStaff As New clsPerson

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SetColours Me
    fillCombos
    dtpLeaveDate.Value = Date
    GetCommonSettings Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub

Private Sub txtLeaveID_Change()
    Set myLeave = New clsLeave
    myLeave.LeaveID = Val(txtLeaveID.text)
    displayDetails
End Sub

Private Sub displayDetails()
    With myLeave
        cmbEmployee.BoundText = .PersonID
        dtpLeaveDate.Value = .LeaveDate
        If .Planned = True Then
            chkPlanned.Value = 1
        Else
            chkPlanned.Value = 0
        End If
        If .HalfDay = True Then
            chkHalfDay.Value = 1
        Else
            chkHalfDay.Value = 0
        End If
        cmbLeaveType.BoundText = .LeaveTypeID
        txtComment.text = .Comments
    End With
End Sub

Private Sub fillCombos()
    Dim Assigners As New clsFillCombo
    Assigners.FillSpecificField cmbEmployee, "Person", "PersonName", True
    Dim allItems As New clsFillCombo
    allItems.FillSpecificFieldOrder cmbLeaveType, "LeaveType", "LeaveTypeName", "LeaveTypeName", True
    
End Sub

Private Sub btnSave_Click()
    If IsNumeric(cmbEmployee.BoundText) = False Then
        MsgBox "Please select an Employee"
        cmbEmployee.SetFocus
        Exit Sub
    End If
    
    With myLeave
        If myLeave.LeaveID = 0 Then
            .AddedDate = Date
            .AddedTime = Time
            .AddedUserID = ProgramVariable.loggedUser.UserID
        End If
        
        .PersonID = Val(cmbEmployee.BoundText)
        
        .Comments = txtComment.text
        .LeaveTypeID = Val(cmbLeaveType.BoundText)
        .LeaveDate = dtpLeaveDate.Value
        
        If chkPlanned.Value = 1 Then
            .Planned = True
        Else
            .Planned = False
        End If
        
        If chkHalfDay.Value = 1 Then
            .HalfDay = True
        Else
            .HalfDay = False
        End If
        .saveData
    End With
    
    Set myLeave = New clsLeave
    
    MsgBox "Saved"
    'Unload Me
    
End Sub
