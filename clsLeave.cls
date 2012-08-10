VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Private varLeaveID As Long
    Private varPersonID As Long
    Private varLeaveDate As Date
    Private varPlanned As Boolean
    Private varHalfDay As Boolean
    Private varLeaveTypeID As Long
    Private varComments As String
    Private varAddedDate As Date
    Private varAddedTime As Date
    Private varAddedUserID As Long
    Private varDeleted As Boolean
    Private varDeletedDate As Date
    Private varDeletedTime As Date
    Private varDeletedUserID As Date

Public Sub saveData()
 
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "SELECT * FROM tblLeave Where LeaveID = " & varLeaveID
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then .AddNew
        !PersonID = varPersonID
        !LeaveDate = varLeaveDate
        !Planned = varPlanned
        !HalfDay = varHalfDay
        !LeaveTypeID = varLeaveTypeID
        !Comments = varComments
        !AddedDate = varAddedDate
        !AddedTime = varAddedTime
        !AddedUserID = varAddedUserID
        !Deleted = varDeleted
        !DeletedDate = varDeletedDate
        !DeletedTime = varDeletedTime
        !DeletedUserID = varDeletedUserID
        .Update
        varLeaveID = !LeaveID
        If .State = 1 Then .Close
    End With
    
End Sub
Public Sub loadData()
 
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "SELECT * FROM tblLeave WHERE LeaveID = " & varLeaveID
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If Not IsNull(!LeaveID) Then
               varLeaveID = !LeaveID
            End If
            If Not IsNull(!PersonID) Then
               varPersonID = !PersonID
            End If
            If Not IsNull(!LeaveDate) Then
               varLeaveDate = !LeaveDate
            End If
            If Not IsNull(!Planned) Then
               varPlanned = !Planned
            End If
            If Not IsNull(!HalfDay) Then
               varHalfDay = !HalfDay
            End If
            If Not IsNull(!LeaveTypeID) Then
               varLeaveTypeID = !LeaveTypeID
            End If
            If Not IsNull(!Comments) Then
               varComments = !Comments
            End If
            If Not IsNull(!AddedDate) Then
               varAddedDate = !AddedDate
            End If
            If Not IsNull(!AddedTime) Then
               varAddedTime = !AddedTime
            End If
            If Not IsNull(!AddedUserID) Then
               varAddedUserID = !AddedUserID
            End If
            If Not IsNull(!Deleted) Then
               varDeleted = !Deleted
            End If
            If Not IsNull(!DeletedDate) Then
               varDeletedDate = !DeletedDate
            End If
            If Not IsNull(!DeletedTime) Then
               varDeletedTime = !DeletedTime
            End If
            If Not IsNull(!DeletedUserID) Then
               varDeletedUserID = !DeletedUserID
            End If
        End If
    If .State = 1 Then .Close
    End With
    
End Sub
Public Sub clearData()
    varLeaveID = 0
    varPersonID = 0
    varLeaveDate = Empty
    varPlanned = False
    varHalfDay = False
    varLeaveTypeID = 0
    varComments = Empty
    varAddedDate = Empty
    varAddedTime = Empty
    varAddedUserID = 0
    varDeleted = False
    varDeletedDate = Empty
    varDeletedTime = Empty
    varDeletedUserID = Empty
End Sub

Public Property Let LeaveID(ByVal vLeaveID As Long)
    Call clearData
    varLeaveID = vLeaveID
    Call loadData
End Property

Public Property Get LeaveID() As Long
    LeaveID = varLeaveID
End Property

Public Property Let PersonID(ByVal vPersonID As Long)
    varPersonID = vPersonID
End Property

Public Property Get PersonID() As Long
    PersonID = varPersonID
End Property

Public Property Let LeaveDate(ByVal vLeaveDate As Date)
    varLeaveDate = vLeaveDate
End Property

Public Property Get LeaveDate() As Date
    LeaveDate = varLeaveDate
End Property

Public Property Let Planned(ByVal vPlanned As Boolean)
    varPlanned = vPlanned
End Property

Public Property Get Planned() As Boolean
    Planned = varPlanned
End Property

Public Property Let HalfDay(ByVal vHalfDay As Boolean)
    varHalfDay = vHalfDay
End Property

Public Property Get HalfDay() As Boolean
    HalfDay = varHalfDay
End Property

Public Property Let LeaveTypeID(ByVal vLeaveTypeID As Long)
    varLeaveTypeID = vLeaveTypeID
End Property

Public Property Get LeaveTypeID() As Long
    LeaveTypeID = varLeaveTypeID
End Property

Public Property Let Comments(ByVal vComments As String)
    varComments = vComments
End Property

Public Property Get Comments() As String
    Comments = varComments
End Property

Public Property Let AddedDate(ByVal vAddedDate As Date)
    varAddedDate = vAddedDate
End Property

Public Property Get AddedDate() As Date
    AddedDate = varAddedDate
End Property

Public Property Let AddedTime(ByVal vAddedTime As Date)
    varAddedTime = vAddedTime
End Property

Public Property Get AddedTime() As Date
    AddedTime = varAddedTime
End Property

Public Property Let AddedUserID(ByVal vAddedUserID As Long)
    varAddedUserID = vAddedUserID
End Property

Public Property Get AddedUserID() As Long
    AddedUserID = varAddedUserID
End Property

Public Property Let Deleted(ByVal vDeleted As Boolean)
    varDeleted = vDeleted
End Property

Public Property Get Deleted() As Boolean
    Deleted = varDeleted
End Property

Public Property Let DeletedDate(ByVal vDeletedDate As Date)
    varDeletedDate = vDeletedDate
End Property

Public Property Get DeletedDate() As Date
    DeletedDate = varDeletedDate
End Property

Public Property Let DeletedTime(ByVal vDeletedTime As Date)
    varDeletedTime = vDeletedTime
End Property

Public Property Get DeletedTime() As Date
    DeletedTime = varDeletedTime
End Property

Public Property Let DeletedUserID(ByVal vDeletedUserID As Date)
    varDeletedUserID = vDeletedUserID
End Property

Public Property Get DeletedUserID() As Date
    DeletedUserID = varDeletedUserID
End Property


