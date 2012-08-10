VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Letter Manager"
   ClientHeight    =   4770
   ClientLeft      =   -150
   ClientTop       =   540
   ClientWidth     =   7515
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileBackUp 
         Caption         =   "Back up"
      End
      Begin VB.Menu mnuFileRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditPersons 
         Caption         =   "Persons"
      End
      Begin VB.Menu mnuEditUsers 
         Caption         =   "Users"
      End
      Begin VB.Menu mnuEditMyDetails 
         Caption         =   "My Details"
      End
      Begin VB.Menu mnuEditLeaveTypes 
         Caption         =   "Leave Types"
      End
   End
   Begin VB.Menu mnuLeaves 
      Caption         =   "Leaves"
      Begin VB.Menu mnuAddLeave 
         Caption         =   "Add Leave"
      End
      Begin VB.Menu mnuLeaveList 
         Caption         =   "Leave List"
      End
      Begin VB.Menu mnuLeaveCharts 
         Caption         =   "Charts"
      End
   End
End
Attribute VB_Name = "MDIMain"
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

Private Sub mnuAddLeave_Click()
    frmLeave.Show
    frmLeave.ZOrder 0
End Sub

Private Sub mnuEditLeaveTypes_Click()
    frmLeaveType.Show
    frmLeaveType.ZOrder 0
End Sub

Private Sub mnuEditMyDetails_Click()
    frmMyDetails.Show
    frmMyDetails.ZOrder 0
End Sub

Private Sub mnuEditPersons_Click()
    frmPerson.Show
    frmPerson.ZOrder 0
End Sub

Private Sub mnuEditPriorities_Click()
    frmLeaveType.Show
    frmLeaveType.ZOrder 0
End Sub




Private Sub mnuEditUsers_Click()
    frmUsers.Show
    frmUsers.ZOrder 0
End Sub

Private Sub mnuFileBackUp_Click()
    frmBackUp.Show
    frmBackUp.ZOrder 0
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileRestore_Click()
    frmRestore.Show
    frmRestore.ZOrder 0
End Sub


Private Sub mnuLeaveCharts_Click()
    frmLeaveChart.Show
    frmLeaveChart.ZOrder 0
End Sub

Private Sub mnuLeaveList_Click()
    frmLeaveList.Show
    frmLeaveList.ZOrder 0
End Sub
