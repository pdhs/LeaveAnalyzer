VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLeaveChart 
   Caption         =   "Leave Chart"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
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
   ScaleHeight     =   9285
   ScaleWidth      =   12450
   Begin VB.CommandButton btnByWeekdays 
      Caption         =   "Unplanned - Pie Chart"
      Height          =   495
      Left            =   9480
      TabIndex        =   7
      Top             =   720
      Width           =   2895
   End
   Begin VB.ListBox lstEmployee 
      Height          =   2760
      Left            =   1800
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   120
      Width           =   7575
   End
   Begin VB.CommandButton btnLeaveTypeBar 
      Caption         =   "Leave Type - Bar Chart"
      Height          =   495
      Left            =   9480
      TabIndex        =   5
      Top             =   120
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   205717507
      CurrentDate     =   40940
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   205717507
      CurrentDate     =   40940
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   1800
      Top             =   3480
      Width           =   7575
   End
   Begin VB.Label Label7 
      Caption         =   "Employee"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "From"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label lblTo 
      Caption         =   "To"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "frmLeaveChart"
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
    Dim temSQL As String

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnByWeekdays_Click()
    Dim excelApp As Excel.Application
    Dim excelWB As Excel.Workbook
    Dim excelWS As Excel.Worksheet
    Dim excelChart As Excel.Chart
    
    Dim rsTem As New ADODB.Recordset
    
    Dim myRow As Integer
    Dim myCol As Integer

    Set excelApp = New Excel.Application
    Set excelWB = excelApp.Workbooks.Add
    Set excelWS = excelWB.Worksheets(1)
    excelApp.Visible = True
    excelApp.UserControl = True

    Dim temTopic As String
    Dim temSubTopic As String
    Dim temSQL As String
    
    
    temTopic = "Leave Type of Staff"
    temSubTopic = "From " & Format(dtpFrom.Value, ProgramVariable.LongDateFormat) & " to " & Format(dtpTo.Value, ProgramVariable.LongDateFormat)

    Dim i As Integer
    

    myRow = 3
    
    excelWS.Cells(myRow, 1).Value = "Name"
    excelWS.Cells(myRow, 2).Value = "Unplanned Leave"
        

        myCol = 2
        With rsTem
            temSQL = "SELECT tblPerson.PersonName, Count(tblLeave.LeaveID) AS CountOfLeaveID " & _
                        "FROM tblLeave LEFT JOIN tblPerson ON tblLeave.PersonID = tblPerson.PersonID " & _
                        "WHERE (((tblLeave.Planned)=False) AND ((tblPerson.Deleted)=False) AND ((tblLeave.Deleted)=False) AND ((tblLeave.LeaveDate) Between #" & Format(dtpFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpTo.Value, "dd MMMM yyyy") & "#)) " & _
                        "GROUP BY tblPerson.PersonName " & _
                        "ORDER BY tblPerson.PersonName"

            If .State = 1 Then .Close
            .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
            While .EOF = False
                If IsNull(!PersonName) = False Then
                    excelWS.Cells(myRow, 1).Value = !PersonName
                End If
                
                If IsNull(!CountOfLeaveID) = False Then
                    excelWS.Cells(myRow, 2).Value = !CountOfLeaveID
                End If
                myRow = myRow + 1
                .MoveNext
            Wend
            .Close
            
        End With

    Set excelChart = excelWB.Charts.Add
    excelChart.ChartType = xlPie
    excelChart.SetSourceData excelWS.Range("A3", "B" & myRow - 1)
    excelChart.Visible = xlSheetVisible
    excelChart.ChartArea.Select
    excelChart.ChartArea.Copy
    Image1.Picture = Clipboard.GetData(vbCFBitmap)
    
    Set excelChart = Nothing
    Set excelWS = Nothing
    Set excelWB = Nothing
    Set excelApp = Nothing
    
    

End Sub

Private Sub btnLeaveTypeBar_Click()
    Dim excelApp As Excel.Application
    Dim excelWB As Excel.Workbook
    Dim excelWS As Excel.Worksheet
    Dim excelChart As Excel.Chart
    
    Dim rsTem As New ADODB.Recordset
    
    Dim myRow As Integer
    Dim myCol As Integer

    Set excelApp = New Excel.Application
    Set excelWB = excelApp.Workbooks.Add
    Set excelWS = excelWB.Worksheets(1)
    excelApp.Visible = True
    excelApp.UserControl = True

    Dim temTopic As String
    Dim temSubTopic As String
    Dim temSQL As String
    
    
    temTopic = "Leave Type of Staff"
    temSubTopic = "From " & Format(dtpFrom.Value, ProgramVariable.LongDateFormat) & " to " & Format(dtpTo.Value, ProgramVariable.LongDateFormat)

    Dim i As Integer
    
    myCol = 2
    myRow = 3
    
    With rsTem
        temSQL = "SELECT tblLeaveType.LeaveTypeName " & _
                    "From tblLeaveType " & _
                    "Where (((tblLeaveType.Deleted) = False)) " & _
                    "ORDER BY tblLeaveType.LeaveTypeName"
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
        While .EOF = False
            excelWS.Cells(myRow, myCol).Value = !LeaveTypeName
            myCol = myCol + 1
            .MoveNext
        Wend
        .Close
        
    End With
    
    myRow = 4
    
    For i = 0 To lstEmployee.ListCount - 1
        If lstEmployee.Selected(i) = True Then
            myCol = 1
            excelWS.Cells(myRow, myCol).Value = lstEmployee.List(i)
            myCol = 2
            With rsTem
                temSQL = "SELECT tblLeaveType.LeaveTypeName, Count(tblLeave.LeaveID) AS CountOfLeaveID " & _
                            "FROM tblLeaveType RIGHT JOIN tblLeave ON tblLeaveType.LeaveTypeID = tblLeave.LeaveTypeID " & _
                            "Where (((tblLeaveType.Deleted) = False) And ((tblLeave.Deleted) = False) And ((tblLeave.PersonID) = " & Val(lstEmployee.ItemData(i)) & ")) " & _
                            "GROUP BY tblLeaveType.LeaveTypeName " & _
                            "ORDER BY tblLeaveType.LeaveTypeName"
                If .State = 1 Then .Close
                .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
                While .EOF = False
                    If IsNull(!CountOfLeaveID) = False Then
                        excelWS.Cells(myRow, myCol).Value = !CountOfLeaveID
                    End If
                    myCol = myCol + 1
                    .MoveNext
                Wend
                .Close
                myRow = myRow + 1
            End With
        End If
    
    Next

    Set excelChart = excelWB.Charts.Add
    excelChart.ChartType = xl3DColumn
    excelChart.SetSourceData excelWS.Range(xl_Col(1) & 3, xl_Col(myCol - 1) & myRow - 1)
    excelChart.Visible = xlSheetVisible
    excelChart.ChartArea.Select
    excelChart.ChartArea.Copy
    Image1.Picture = Clipboard.GetData(vbCFBitmap)
    
    Set excelChart = Nothing
    Set excelWS = Nothing
    Set excelWB = Nothing
    Set excelApp = Nothing
    
    
End Sub


Function xl_Col(ByRef Col_No) As String
'returns Excel column name from numeric position (e.g.: col_no 27 returns "AA")
'by Si_the_geek (VBForums.com)
 
                                      'Only allow valid columns
  If Col_No < 1 Or Col_No > 256 Then Exit Function
 
  If Col_No < 27 Then                  'Single letter
    xl_Col = Chr(Col_No + 64)
  Else                                 'Two letters
    xl_Col = Chr(Int((Col_No - 1) / 26) + 64) & _
             Chr(((Col_No - 1) Mod 26) + 1 + 64)
  End If
 
End Function

Function xl_ColNo(Col_Name) As Integer
'returns an Excel column number from its name (e.g.: col_name "AA" returns  27)
'by Si_the_geek (VBForums.com)
 
  Col_Name = UCase(Trim(Col_Name))
  Select Case Len(Col_Name)
  Case 1:     xl_ColNo = Asc(Col_Name) - 64
  Case 2:     xl_ColNo = ((Asc(Left(Col_Name, 1)) - 64) * 26) _
                       + (Asc(Right(Col_Name, 1)) - 64)
  End Select
 
End Function

Private Sub Form_Load()
    SetColours Me
    fillCombos
    dtpFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    dtpTo.Value = Date
    GetCommonSettings Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub

Private Sub fillCombos()
    Dim rsTem As New ADODB.Recordset
    temSQL = "SELECT tblPerson.PersonName, tblPerson.PersonID " & _
                "From tblPerson " & _
                "Where (((tblPerson.Deleted) = False)) " & _
                "ORDER BY tblPerson.PersonName"
    lstEmployee.Clear
    Dim i As Integer
    With rsTem
        i = 0
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
        While .EOF = False
            lstEmployee.AddItem !PersonName
            lstEmployee.ItemData(i) = !PersonID
            i = i + 1
            .MoveNext
        Wend
        .Close
    End With
End Sub
