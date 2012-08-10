VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLetter"
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
    Dim temSQL As String
    Private varLetterID As Long
    Private varLetterNumber As String
    Private varLetterDate As Date
    Private varLetterTopic As String
    Private varLetterContent As String
    Private varLetterComments As String
    Private varSenderUnitID As Long
    Private varSenderPersonID As Long
    Private varReceiverUnitID As Long
    Private varReceiverPersonID As Long
    Private varReceivedDate As Date
    Private varReceivedTime As Date
    Private varPriorityID As Long
    Private varNeedReply As Boolean
    Private varReplyDueDate As Date
    Private varReplied As Boolean
    Private varRepliedDate As Date
    Private varRepliedTime As Date
    Private varReplyComments As String
    Private varDeleted As Boolean
    Private varDeletedDate As Date
    Private varDeletedTime As Date
    Private varDeletedUserID As Date
    Private varAddedDate As Date
    Private varAddedTime As Date
    Private varAddedUserID As Long
    Private varAssigned As Boolean
    Private varAssignedPersonID As Long
    Private varAssignedDate As Date
    Private varAssignedTime As Date
    Private varAssignedComments As String
    Private varNeedComplete As Boolean
    Private varCompleted As Boolean
    Private varCompleteDueDate As Date
    Private varCompletedDate As Date
    Private varCompletedTime As Date
    Private varCompletedComments As String

Public Sub saveData()
 
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "SELECT * FROM tblLetter Where LetterID = " & varLetterID
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then .AddNew
        !LetterNumber = varLetterNumber
        !LetterDate = varLetterDate
        !LetterTopic = varLetterTopic
        !LetterContent = varLetterContent
        !LetterComments = varLetterComments
        !SenderUnitID = varSenderUnitID
        !SenderPersonID = varSenderPersonID
        !ReceiverUnitID = varReceiverUnitID
        !ReceiverPersonID = varReceiverPersonID
        !ReceivedDate = varReceivedDate
        !ReceivedTime = varReceivedTime
        !PriorityID = varPriorityID
        !NeedReply = varNeedReply
        !ReplyDueDate = varReplyDueDate
        !Replied = varReplied
        !RepliedDate = varRepliedDate
        !RepliedTime = varRepliedTime
        !ReplyComments = varReplyComments
        !Deleted = varDeleted
        !DeletedDate = varDeletedDate
        !DeletedTime = varDeletedTime
        !DeletedUserID = varDeletedUserID
        !AddedDate = varAddedDate
        !AddedTime = varAddedTime
        !AddedUserID = varAddedUserID
        !Assigned = varAssigned
        !AssignedPersonID = varAssignedPersonID
        !AssignedDate = varAssignedDate
        !AssignedTime = varAssignedTime
        !AssignedComments = varAssignedComments
        !NeedComplete = varNeedComplete
        !Completed = varCompleted
        !CompleteDueDate = varCompleteDueDate
        !CompletedDate = varCompletedDate
        !CompletedTime = varCompletedTime
        !CompletedComments = varCompletedComments
        .Update
        varLetterID = !LetterID
        If .State = 1 Then .Close
    End With
    
End Sub
Public Sub loadData()
 
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSQL = "SELECT * FROM tblLetter WHERE LetterID = " & varLetterID
        If .State = 1 Then .Close
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If Not IsNull(!LetterID) Then
               varLetterID = !LetterID
            End If
            If Not IsNull(!LetterNumber) Then
               varLetterNumber = !LetterNumber
            End If
            If Not IsNull(!LetterDate) Then
               varLetterDate = !LetterDate
            End If
            If Not IsNull(!LetterTopic) Then
               varLetterTopic = !LetterTopic
            End If
            If Not IsNull(!LetterContent) Then
               varLetterContent = !LetterContent
            End If
            If Not IsNull(!LetterComments) Then
               varLetterComments = !LetterComments
            End If
            If Not IsNull(!SenderUnitID) Then
               varSenderUnitID = !SenderUnitID
            End If
            If Not IsNull(!SenderPersonID) Then
               varSenderPersonID = !SenderPersonID
            End If
            If Not IsNull(!ReceiverUnitID) Then
               varReceiverUnitID = !ReceiverUnitID
            End If
            If Not IsNull(!ReceiverPersonID) Then
               varReceiverPersonID = !ReceiverPersonID
            End If
            If Not IsNull(!ReceivedDate) Then
               varReceivedDate = !ReceivedDate
            End If
            If Not IsNull(!ReceivedTime) Then
               varReceivedTime = !ReceivedTime
            End If
            If Not IsNull(!PriorityID) Then
               varPriorityID = !PriorityID
            End If
            If Not IsNull(!NeedReply) Then
               varNeedReply = !NeedReply
            End If
            If Not IsNull(!ReplyDueDate) Then
               varReplyDueDate = !ReplyDueDate
            End If
            If Not IsNull(!Replied) Then
               varReplied = !Replied
            End If
            If Not IsNull(!RepliedDate) Then
               varRepliedDate = !RepliedDate
            End If
            If Not IsNull(!RepliedTime) Then
               varRepliedTime = !RepliedTime
            End If
            If Not IsNull(!ReplyComments) Then
               varReplyComments = !ReplyComments
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
            If Not IsNull(!Assigned) Then
               varAssigned = !Assigned
            End If
            If Not IsNull(!AssignedPersonID) Then
               varAssignedPersonID = !AssignedPersonID
            End If
            If Not IsNull(!AssignedDate) Then
               varAssignedDate = !AssignedDate
            End If
            If Not IsNull(!AssignedTime) Then
               varAssignedTime = !AssignedTime
            End If
            If Not IsNull(!AssignedComments) Then
               varAssignedComments = !AssignedComments
            End If
            If Not IsNull(!NeedComplete) Then
               varNeedComplete = !NeedComplete
            End If
            If Not IsNull(!Completed) Then
               varCompleted = !Completed
            End If
            If Not IsNull(!CompleteDueDate) Then
               varCompleteDueDate = !CompleteDueDate
            End If
            If Not IsNull(!CompletedDate) Then
               varCompletedDate = !CompletedDate
            End If
            If Not IsNull(!CompletedTime) Then
               varCompletedTime = !CompletedTime
            End If
            If Not IsNull(!CompletedComments) Then
               varCompletedComments = !CompletedComments
            End If
        End If
    If .State = 1 Then .Close
    End With
    
End Sub
Public Sub clearData()
    varLetterID = 0
    varLetterNumber = Empty
    varLetterDate = Empty
    varLetterTopic = Empty
    varLetterContent = Empty
    varLetterComments = Empty
    varSenderUnitID = 0
    varSenderPersonID = 0
    varReceiverUnitID = 0
    varReceiverPersonID = 0
    varReceivedDate = Empty
    varReceivedTime = Empty
    varPriorityID = 0
    varNeedReply = False
    varReplyDueDate = Empty
    varReplied = False
    varRepliedDate = Empty
    varRepliedTime = Empty
    varReplyComments = Empty
    varDeleted = False
    varDeletedDate = Empty
    varDeletedTime = Empty
    varDeletedUserID = Empty
    varAddedDate = Empty
    varAddedTime = Empty
    varAddedUserID = 0
    varAssigned = False
    varAssignedPersonID = 0
    varAssignedDate = Empty
    varAssignedTime = Empty
    varAssignedComments = Empty
    varNeedComplete = False
    varCompleted = False
    varCompleteDueDate = Empty
    varCompletedDate = Empty
    varCompletedTime = Empty
    varCompletedComments = Empty
End Sub

Public Property Let LetterID(ByVal vLetterID As Long)
    Call clearData
    varLetterID = vLetterID
    Call loadData
End Property

Public Property Get LetterID() As Long
    LetterID = varLetterID
End Property

Public Property Let LetterNumber(ByVal vLetterNumber As String)
    varLetterNumber = vLetterNumber
End Property

Public Property Get LetterNumber() As String
    LetterNumber = varLetterNumber
End Property

Public Property Let LetterDate(ByVal vLetterDate As Date)
    varLetterDate = vLetterDate
End Property

Public Property Get LetterDate() As Date
    LetterDate = varLetterDate
End Property

Public Property Let LetterTopic(ByVal vLetterTopic As String)
    varLetterTopic = vLetterTopic
End Property

Public Property Get LetterTopic() As String
    LetterTopic = varLetterTopic
End Property

Public Property Let LetterContent(ByVal vLetterContent As String)
    varLetterContent = vLetterContent
End Property

Public Property Get LetterContent() As String
    LetterContent = varLetterContent
End Property

Public Property Let LetterComments(ByVal vLetterComments As String)
    varLetterComments = vLetterComments
End Property

Public Property Get LetterComments() As String
    LetterComments = varLetterComments
End Property

Public Property Let SenderUnitID(ByVal vSenderUnitID As Long)
    varSenderUnitID = vSenderUnitID
End Property

Public Property Get SenderUnitID() As Long
    SenderUnitID = varSenderUnitID
End Property

Public Property Let SenderPersonID(ByVal vSenderPersonID As Long)
    varSenderPersonID = vSenderPersonID
End Property

Public Property Get SenderPersonID() As Long
    SenderPersonID = varSenderPersonID
End Property

Public Property Let ReceiverUnitID(ByVal vReceiverUnitID As Long)
    varReceiverUnitID = vReceiverUnitID
End Property

Public Property Get ReceiverUnitID() As Long
    ReceiverUnitID = varReceiverUnitID
End Property

Public Property Let ReceiverPersonID(ByVal vReceiverPersonID As Long)
    varReceiverPersonID = vReceiverPersonID
End Property

Public Property Get ReceiverPersonID() As Long
    ReceiverPersonID = varReceiverPersonID
End Property

Public Property Let ReceivedDate(ByVal vReceivedDate As Date)
    varReceivedDate = vReceivedDate
End Property

Public Property Get ReceivedDate() As Date
    ReceivedDate = varReceivedDate
End Property

Public Property Let ReceivedTime(ByVal vReceivedTime As Date)
    varReceivedTime = vReceivedTime
End Property

Public Property Get ReceivedTime() As Date
    ReceivedTime = varReceivedTime
End Property

Public Property Let PriorityID(ByVal vPriorityID As Long)
    varPriorityID = vPriorityID
End Property

Public Property Get PriorityID() As Long
    PriorityID = varPriorityID
End Property

Public Property Let NeedReply(ByVal vNeedReply As Boolean)
    varNeedReply = vNeedReply
End Property

Public Property Get NeedReply() As Boolean
    NeedReply = varNeedReply
End Property

Public Property Let ReplyDueDate(ByVal vReplyDueDate As Date)
    varReplyDueDate = vReplyDueDate
End Property

Public Property Get ReplyDueDate() As Date
    ReplyDueDate = varReplyDueDate
End Property

Public Property Let Replied(ByVal vReplied As Boolean)
    varReplied = vReplied
End Property

Public Property Get Replied() As Boolean
    Replied = varReplied
End Property

Public Property Let RepliedDate(ByVal vRepliedDate As Date)
    varRepliedDate = vRepliedDate
End Property

Public Property Get RepliedDate() As Date
    RepliedDate = varRepliedDate
End Property

Public Property Let RepliedTime(ByVal vRepliedTime As Date)
    varRepliedTime = vRepliedTime
End Property

Public Property Get RepliedTime() As Date
    RepliedTime = varRepliedTime
End Property

Public Property Let ReplyComments(ByVal vReplyComments As String)
    varReplyComments = vReplyComments
End Property

Public Property Get ReplyComments() As String
    ReplyComments = varReplyComments
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

Public Property Let Assigned(ByVal vAssigned As Boolean)
    varAssigned = vAssigned
End Property

Public Property Get Assigned() As Boolean
    Assigned = varAssigned
End Property

Public Property Let AssignedPersonID(ByVal vAssignedPersonID As Long)
    varAssignedPersonID = vAssignedPersonID
End Property

Public Property Get AssignedPersonID() As Long
    AssignedPersonID = varAssignedPersonID
End Property

Public Property Let AssignedDate(ByVal vAssignedDate As Date)
    varAssignedDate = vAssignedDate
End Property

Public Property Get AssignedDate() As Date
    AssignedDate = varAssignedDate
End Property

Public Property Let AssignedTime(ByVal vAssignedTime As Date)
    varAssignedTime = vAssignedTime
End Property

Public Property Get AssignedTime() As Date
    AssignedTime = varAssignedTime
End Property

Public Property Let AssignedComments(ByVal vAssignedComments As String)
    varAssignedComments = vAssignedComments
End Property

Public Property Get AssignedComments() As String
    AssignedComments = varAssignedComments
End Property

Public Property Let NeedComplete(ByVal vNeedComplete As Boolean)
    varNeedComplete = vNeedComplete
End Property

Public Property Get NeedComplete() As Boolean
    NeedComplete = varNeedComplete
End Property

Public Property Let Completed(ByVal vCompleted As Boolean)
    varCompleted = vCompleted
End Property

Public Property Get Completed() As Boolean
    Completed = varCompleted
End Property

Public Property Let CompleteDueDate(ByVal vCompleteDueDate As Date)
    varCompleteDueDate = vCompleteDueDate
End Property

Public Property Get CompleteDueDate() As Date
    CompleteDueDate = varCompleteDueDate
End Property

Public Property Let CompletedDate(ByVal vCompletedDate As Date)
    varCompletedDate = vCompletedDate
End Property

Public Property Get CompletedDate() As Date
    CompletedDate = varCompletedDate
End Property

Public Property Let CompletedTime(ByVal vCompletedTime As Date)
    varCompletedTime = vCompletedTime
End Property

Public Property Get CompletedTime() As Date
    CompletedTime = varCompletedTime
End Property

Public Property Let CompletedComments(ByVal vCompletedComments As String)
    varCompletedComments = vCompletedComments
End Property

Public Property Get CompletedComments() As String
    CompletedComments = varCompletedComments
End Property


