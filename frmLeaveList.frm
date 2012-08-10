VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmLeaveList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Letters"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
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
   ScaleHeight     =   8580
   ScaleWidth      =   15270
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7080
      Top             =   4080
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "C&lose"
      Height          =   495
      Left            =   13920
      TabIndex        =   1
      Top             =   8040
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid gridList 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   11033
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   16384003
      CurrentDate     =   40940
   End
   Begin MSDataListLib.DataCombo cmbEmployee 
      Height          =   360
      Left            =   1800
      TabIndex        =   3
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
      TabIndex        =   4
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   16384003
      CurrentDate     =   40940
   End
   Begin VB.Label lblTo 
      Caption         =   "To"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Type"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Employee"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmLeaveList"
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

    
    Dim formActive As Boolean
    Dim activeCount As Integer

Private Sub btnSave_Click()
    Unload Me
End Sub

Private Sub cmbEmployee_Change()
    Call process
End Sub


Private Sub cmbEmployee_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbEmployee.text = Empty
    End If
End Sub

Private Sub cmbLeaveType_Change()
    Call process
End Sub


Private Sub cmbLeaveType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbLeaveType.text = Empty
    End If
End Sub

Private Sub dtpFrom_Change()
    Call process
End Sub



Private Sub dtpTo_Change()
    Call process
End Sub

Private Sub Form_Load()
    Call prepareResize
    SetColours Me
    Me.BorderStyle = 2
    Call fillCombos
    GetCommonSettings Me
    dtpFrom.Value = DateSerial(Year(Date), 1, 1)
    dtpTo.Value = Date
    Call process
End Sub

Private Sub fillCombos()
    Dim Assigners As New clsFillCombo
    Assigners.FillSpecificField cmbEmployee, "Person", "PersonName", True
    Dim allItems As New clsFillCombo
    allItems.FillSpecificFieldOrder cmbLeaveType, "LeaveType", "LeaveTypeName", "LeaveTypeName", True
    
End Sub


Private Sub Form_Activate()
    formActive = True
    activeCount = 0
End Sub

Private Sub Form_Deactivate()
    formActive = False
    activeCount = 0
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


Private Sub gridList_DblClick()
    Dim myForm As New frmLeave
    myForm.Show
    myForm.ZOrder 0
    myForm.txtLeaveID = Val(gridList.TextMatrix(gridList.row, 0))
End Sub


Private Sub process()
    Dim temTopic As String
    Dim temSubTopic As String
    Dim temSelect As String
    Dim temFrom As String
    Dim temWhere As String
    Dim temGroupBy As String
    Dim temOrderBy As String
    Dim temSQL As String
    
    
    Dim D(0) As Integer
    Dim p(0) As Integer
    
    temTopic = "Leave List"
    
    
    
    

    temWhere = "WHERE ((tblLeave.LeaveDate) Between #" & Format(dtpFrom.Value, ProgramVariable.LongDateFormat) & "#  And #" & Format(dtpTo.Value, ProgramVariable.LongDateFormat) & "# "
    
    If IsNumeric(cmbEmployee.BoundText) = True Then
        temWhere = temWhere & " AND ((tblLeave.PersonID)=" & cmbEmployee.BoundText & ")"
    End If
    
    If IsNumeric(cmbLeaveType.BoundText) = True Then
        temWhere = temWhere & " AND ((tblLeave.LeaveTypeID)=" & cmbLeaveType.BoundText & ")"
    End If
    
    temWhere = temWhere & ") "
    
    temSelect = "SELECT tblLeave.LeaveID, tblPerson.PersonName AS [Employee Name], Format$(tblLeave.LeaveDate,'yyyy mmmm dd') AS [Leave Date], tblLeaveType.LeaveTypeName AS [Leave Type], FORMAT$((tblLeave.LeaveDate), 'dddd') AS [Day of Week], Format$(tblLeave.Planned,'Yes/no') AS [Planned], Format$(tblLeave.HalfDay,'Yes/no') AS [Half-day], tblLeave.Comments "
    temFrom = "FROM (tblLeave LEFT JOIN tblPerson ON tblLeave.PersonID = tblPerson.PersonID) LEFT JOIN tblLeaveType ON tblLeave.LeaveTypeID = tblLeaveType.LeaveTypeID "
    
    
    temGroupBy = ""
    
    temSQL = temSelect & temFrom & temWhere & temGroupBy & temOrderBy
    FillAnyGrid temSQL, gridList, 0, D, p
    gridList.ColWidth(0) = 0
    
End Sub


