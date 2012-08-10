VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFillCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Auther : Dr. M. H. B. Ariyaratne
'          buddhika.ari@gmail.com
'          buddhika_ari@yahoo.com
'          +94 71 58 12399
'          GPL Licence

Option Explicit
    Dim rsFill As New ADODB.Recordset
    Dim temSQL As String
    
    
    
Public Sub FillAnyCombo(ComboToFill As DataCombo, table As String, Optional DoNotIncludeDeleted As Boolean)
    temSQL = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then temSQL = temSQL & " Where Deleted = False "
    temSQL = temSQL & " Order by " & table
    With rsFill
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = table
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillOrderCombo(ComboToFill As DataCombo, table As String, OrderField As String, Optional DoNotIncludeDeleted As Boolean)
    temSQL = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then temSQL = temSQL & " Where Deleted = False "
    temSQL = temSQL & " Order by " & OrderField
    With rsFill
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = table
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillBoolCombo(ComboToFill As DataCombo, table As String, ListField As String, BoolField As String, Optional DoNotIncludeDeleted As Boolean)
    temSQL = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then
        temSQL = temSQL & " Where Deleted = False AND " & BoolField & " = True "
    Else
        temSQL = temSQL & " Where " & BoolField & " = True "
    End If
    temSQL = temSQL & " Order by " & ListField
    With rsFill
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = ListField
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillSpecificFieldOrder(ComboToFill As DataCombo, table As String, DisplayField As String, OrderField As String, Optional DoNotIncludeDeleted As Boolean)
    temSQL = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then temSQL = temSQL & " Where Deleted = False "
    temSQL = temSQL & " Order by " & OrderField
    With rsFill
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = DisplayField
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillSpecificField(ComboToFill As DataCombo, table As String, DisplayField As String, Optional DoNotIncludeDeleted As Boolean)
    temSQL = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then temSQL = temSQL & " Where Deleted = False "
    temSQL = temSQL & " Order by " & DisplayField
    With rsFill
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = DisplayField
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillSpecificFieldBoolCombo(ComboToFill As DataCombo, table As String, ListField As String, DisplayField As String, BoolField As String, Optional DoNotIncludeDeleted As Boolean)
    temSQL = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then
        temSQL = temSQL & " Where Deleted = False AND " & BoolField & " = True "
    Else
        temSQL = temSQL & " Where " & BoolField & " = True "
    End If
    temSQL = temSQL & " Order by " & DisplayField
    With rsFill
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = DisplayField
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillLongCombo(ComboToFill As DataCombo, table As String, ListField As String, LongField As String, LongValue As Long, Optional DoNotIncludeDeleted As Boolean)
    temSQL = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then
        temSQL = temSQL & " Where Deleted = False AND " & LongField & " = " & LongValue
    Else
        temSQL = temSQL & " Where " & LongField & " =  " & LongValue
    End If
    temSQL = temSQL & " Order by " & ListField
    With rsFill
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = ListField
        .BoundColumn = table & "ID"
    End With
End Sub

Public Sub FillLongBoolCombo(ComboToFill As DataCombo, table As String, ListField As String, LongField As String, LongValue As Long, BoolField As String, Optional DoNotIncludeDeleted As Boolean)
    temSQL = "Select * from tbl" & table
    If DoNotIncludeDeleted = True Then
        temSQL = temSQL & " Where Deleted = False AND " & LongField & " = " & LongValue & " AND " & BoolField & " = True "
    Else
        temSQL = temSQL & " Where " & LongField & " =  " & LongValue & " AND " & BoolField & " = True "
    End If
    temSQL = temSQL & " Order by " & ListField
    With rsFill
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
    End With
    With ComboToFill
        Set .RowSource = rsFill
        .ListField = ListField
        .BoundColumn = table & "ID"
    End With
End Sub



