VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Auther : Dr. M. H. B. Ariyaratne
'          buddhika.ari@gmail.com
'          buddhika_ari@yahoo.com
'          +94 71 58 12399
'          GPL Licence

Option Explicit
    Dim temSQL As String
    Private varPersonID As Long
    Private varPersonName As String
    Private varDeleted As Boolean
    Private varDeletedDate As Date
    Private varDeletedTime As Date
    Private varDeletedUserID As Date
    Private varAddedDate As Date
    Private varAddedTime As Date
    Private varAddedUserID As Long

Public Sub saveData()
 
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "SELECT * FROM tblPerson Where PersonID = " & varPersonID
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then .AddNew
        !PersonName = varPersonName
        !Deleted = varDeleted
        !DeletedDate = varDeletedDate
        !DeletedTime = varDeletedTime
        !DeletedUserID = varDeletedUserID
        !AddedDate = varAddedDate
        !AddedTime = varAddedTime
        !AddedUserID = varAddedUserID
        .Update
        varPersonID = !PersonID
        If .State = 1 Then .Close
    End With
    
End Sub
Public Sub loadData()
 
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "SELECT * FROM tblPerson WHERE PersonID = " & varPersonID
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If Not IsNull(!PersonID) Then
               varPersonID = !PersonID
            End If
            If Not IsNull(!PersonName) Then
               varPersonName = !PersonName
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
            If Not IsNull(!AddedDate) Then
               varAddedDate = !AddedDate
            End If
            If Not IsNull(!AddedTime) Then
               varAddedTime = !AddedTime
            End If
            If Not IsNull(!AddedUserID) Then
               varAddedUserID = !AddedUserID
            End If
        End If
    If .State = 1 Then .Close
    End With
    
End Sub
Public Sub clearData()
    varPersonID = 0
    varPersonName = Empty
    varDeleted = False
    varDeletedDate = Empty
    varDeletedTime = Empty
    varDeletedUserID = Empty
    varAddedDate = Empty
    varAddedTime = Empty
    varAddedUserID = 0
End Sub

Public Property Let PersonID(ByVal vPersonID As Long)
    Call clearData
    varPersonID = vPersonID
    Call loadData
End Property

Public Property Get PersonID() As Long
    PersonID = varPersonID
End Property

Public Property Let PersonName(ByVal vPersonName As String)
    varPersonName = vPersonName
End Property

Public Property Get PersonName() As String
    PersonName = varPersonName
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


