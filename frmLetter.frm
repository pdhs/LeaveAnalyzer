VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmLetter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Letter"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12435
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
   ScaleHeight     =   8760
   ScaleWidth      =   12435
   Begin VB.Timer Timer1 
      Left            =   8760
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C&lose"
      Height          =   495
      Left            =   11040
      TabIndex        =   39
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox txtLetterID 
      Height          =   375
      Left            =   6840
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Replied"
      Height          =   1215
      Left            =   6240
      TabIndex        =   31
      Top             =   6720
      Width           =   6015
      Begin VB.TextBox txtReplyComments 
         Height          =   375
         Left            =   1800
         TabIndex        =   33
         Top             =   720
         Width           =   4095
      End
      Begin VB.CheckBox chkReplied 
         Caption         =   "Replied"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpReplied 
         Height          =   375
         Left            =   1800
         TabIndex        =   34
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   16449539
         CurrentDate     =   40940
      End
      Begin VB.Label Label2 
         Caption         =   "Comments"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Complete"
      Height          =   1215
      Left            =   6240
      TabIndex        =   25
      Top             =   5520
      Width           =   6015
      Begin VB.CheckBox chkCompleted 
         Caption         =   "Completed"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtCompleteComments 
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         Top             =   720
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpComplete 
         Height          =   375
         Left            =   1800
         TabIndex        =   28
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   16449539
         CurrentDate     =   40940
      End
      Begin VB.Label Label4 
         Caption         =   "Comments"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Assign"
      Height          =   2415
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Width           =   6015
      Begin VB.CheckBox chkAssigned 
         Caption         =   "Assigned"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtAssignComments 
         Height          =   615
         Left            =   1800
         TabIndex        =   23
         Top             =   1200
         Width           =   4095
      End
      Begin VB.CheckBox chkCompleteBy 
         Caption         =   "Complete By"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpAssigned 
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   16449539
         CurrentDate     =   40940
      End
      Begin MSDataListLib.DataCombo cmbAssigner 
         Height          =   360
         Left            =   1800
         TabIndex        =   20
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpCompleteBy 
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Top             =   1920
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   16449539
         CurrentDate     =   40940
      End
      Begin VB.Label lblNewAssigner 
         Height          =   375
         Left            =   4800
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Comments"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Assigned To"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.CheckBox chkReply 
      Caption         =   "Reply By"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
   End
   Begin VB.TextBox txtContent 
      Height          =   1935
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   7575
   End
   Begin VB.TextBox txtLetterNo 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   7575
   End
   Begin VB.TextBox txtTopic 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   7575
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   9720
      TabIndex        =   0
      Top             =   8040
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpReplyBefore 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   5040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   16449539
      CurrentDate     =   40940
   End
   Begin MSComCtl2.DTPicker dtpLetterDate 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   16449539
      CurrentDate     =   40940
   End
   Begin MSComCtl2.DTPicker dtpReceivedDate 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   16449539
      CurrentDate     =   40940
   End
   Begin MSDataListLib.DataCombo cmbPriority 
      Height          =   360
      Left            =   1800
      TabIndex        =   8
      Top             =   4560
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSender 
      Height          =   360
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label lblNewFrom 
      Height          =   375
      Left            =   9480
      TabIndex        =   36
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Letter Date"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Received Date"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "Number"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "Topic"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Content"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "Priority"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   2055
   End
End
Attribute VB_Name = "frmLetter"
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
    Dim myLetter As New clsLetter
    Dim mySender As New clsPerson
    Dim myAssigner As New clsPerson

    
Private Sub cmbSender_LostFocus()
    If IsNumeric(cmbSender.BoundText) = False Then
        With mySender
            .AddedDate = Date
            .AddedTime = Time
            .AddedUserID = ProgramVariable.loggedUser.UserID
            .PersonName = cmbSender.text
        End With
        lblNewFrom.Caption = "New"
    Else
        mySender.PersonID = Val(cmbSender.BoundText)
        lblNewFrom.Caption = ""
    End If
End Sub

Private Sub cmbAssigner_LostFocus()
    If IsNumeric(cmbAssigner.BoundText) = False Then
        With myAssigner
            .AddedDate = Date
            .AddedTime = Time
            .AddedUserID = ProgramVariable.loggedUser.UserID
            .PersonName = cmbAssigner.text
        End With
        lblNewAssigner.Caption = "New"
    Else
        myAssigner.PersonID = Val(cmbAssigner.BoundText)
        lblNewAssigner.Caption = ""
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    SetColours Me
    fillCombos
    dtpLetterDate.Value = Date
    dtpAssigned.Value = Date
    dtpComplete.Value = Date
    dtpCompleteBy.Value = Date
    dtpReceivedDate.Value = Date
    dtpReplied.Value = Date
    dtpReplyBefore.Value = Date
    
    GetCommonSettings Me
End Sub



Private Sub txtLetterID_Change()
    Set myLetter = New clsLetter
    myLetter.LetterID = Val(txtLetterID.text)
    displaydetails
End Sub

Private Sub displaydetails()
    With myLetter
        

    
        mySender.PersonID = .SenderPersonID
        myAssigner.PersonID = .AssignedPersonID
        
        cmbSender.BoundText = .SenderPersonID
        
        txtContent.text = .LetterContent
        txtLetterNo.text = .LetterNumber
        txtTopic.text = .LetterTopic
        dtpLetterDate.Value = .LetterDate
        
        If .NeedReply = True Then
            chkReply.Value = 1
            dtpReplyBefore.Value = .ReplyDueDate
        Else
            chkReply.Value = 0
        End If
        
        cmbPriority.BoundText = .PriorityID
        dtpReceivedDate.Value = .ReceivedDate
        
        
        If .Assigned = True Then
            chkAssigned.Value = 1
            txtAssignComments.text = .AssignedComments
            dtpAssigned.Value = .AssignedDate
            cmbAssigner.BoundText = .AssignedPersonID
        Else
            chkAssigned.Value = 0
            .AssignedComments = txtAssignComments.text
            cmbAssigner.text = Empty
        End If
        
        If .NeedComplete = True Then
            chkCompleteBy.Value = 1
        Else
            chkCompleteBy.Value = 0
        End If
        
        If .Completed = True Then
            chkCompleted.Value = 1
            txtCompleteComments.text = .CompletedComments
        Else
            chkCompleted.Value = 0
            txtCompleteComments.text = .CompletedComments
        End If
        
        If .Replied = True Then
             chkReplied.Value = 1
            dtpReplied.Value = .RepliedDate
            txtReplyComments.text = .ReplyComments
        Else
             chkReplied.Value = 0
            txtReplyComments.text = .ReplyComments
        End If
    End With
End Sub

Private Sub fillCombos()
    Dim Assigners As New clsFillCombo
    Assigners.FillSpecificField cmbAssigner, "Person", "PersonName", True
    Dim Sender As New clsFillCombo
    Sender.FillSpecificField cmbSender, "Person", "PersonName", True
    Dim Priority As New clsFillCombo
    Priority.FillSpecificFieldOrder cmbPriority, "Priority", "PriorityName", "OrderNo", True
End Sub

Private Sub btnSave_Click()

    
    With myLetter
        If myLetter.LetterID = 0 Then
            If Trim(mySender.PersonName) <> "" Then mySender.saveData
            If Trim(myAssigner.PersonName) <> "" Then myAssigner.saveData
            .AddedDate = Date
            .AddedTime = Time
            .AddedUserID = ProgramVariable.loggedUser.UserID
        End If
        
        .SenderPersonID = mySender.PersonID
        
        .LetterContent = txtContent.text
        .LetterNumber = txtLetterNo.text
        .LetterTopic = txtTopic.text
        .LetterDate = dtpLetterDate.Value
        
        If chkReply.Value = 1 Then
            .NeedReply = True
            .ReplyDueDate = dtpReplyBefore.Value
        Else
            .NeedReply = False
            .ReplyDueDate = Empty
        End If
        
        .PriorityID = Val(cmbPriority.BoundText)
        .ReceivedDate = dtpReceivedDate.Value
        
        
        If chkAssigned.Value = 1 Then
            .Assigned = True
            .AssignedComments = txtAssignComments.text
            .AssignedDate = dtpAssigned.Value
            .AssignedPersonID = myAssigner.PersonID
        Else
            .Assigned = False
            .AssignedComments = txtAssignComments.text
            .AssignedDate = Empty
            .AssignedPersonID = Empty
        End If
        
        If chkCompleteBy.Value = 1 Then
            .CompleteDueDate = dtpCompleteBy.Value
            .NeedComplete = True
        Else
            .CompleteDueDate = Empty
            .NeedComplete = False
        End If
        
        If chkCompleted.Value = 1 Then
            .Completed = True
            .CompletedComments = txtCompleteComments.text
            .CompletedDate = dtpComplete.Value
        Else
            .Completed = False
            .CompletedComments = txtCompleteComments.text
            .CompletedDate = Empty
        End If
        
        If chkReplied.Value = 1 Then
            .Replied = True
            .RepliedDate = dtpReplied.Value
            .ReplyComments = txtReplyComments.text
        Else
            .Replied = False
            .RepliedDate = Empty
            .ReplyComments = txtReplyComments.text
        End If
        
        .saveData
    End With
    
    MsgBox "Saved"
    Unload Me
    
End Sub
