VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Private varVersionID As Long
    Private varVersionNo As Long

Public Sub saveData()
 
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "SELECT * FROM tblVersion Where VersionID = " & varVersionID
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then .AddNew
        !VersionNo = varVersionNo
        .Update
        varVersionID = !VersionID
        If .State = 1 Then .Close
    End With
    
End Sub
Public Sub loadData()
 
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "SELECT * FROM tblVersion ORDER BY VersionID"
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If Not IsNull(!VersionID) Then
               varVersionID = !VersionID
            End If
            If Not IsNull(!VersionNo) Then
               varVersionNo = !VersionNo
            End If
        End If
    If .State = 1 Then .Close
    End With
    
End Sub
Public Sub clearData()
    varVersionID = 0
    varVersionNo = 0
End Sub

Public Property Let VersionID(ByVal vVersionID As Long)
    Call clearData
    varVersionID = vVersionID
    Call loadData
End Property

Public Property Get VersionID() As Long
    VersionID = varVersionID
End Property

Public Property Let VersionNo(ByVal vVersionNo As Long)
    varVersionNo = vVersionNo
End Property

Public Property Get VersionNo() As Long
    VersionNo = varVersionNo
End Property
