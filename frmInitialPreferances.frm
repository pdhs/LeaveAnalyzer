VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmInitialPreferances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Database"
   ClientHeight    =   1500
   ClientLeft      =   4440
   ClientTop       =   1680
   ClientWidth     =   7350
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
   ScaleHeight     =   1500
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton bttnClose 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton bttnExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton bttnSelectDatabasePath 
      Caption         =   "&Select Path"
      Height          =   375
      Left            =   120
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.Frame Frame21 
      Caption         =   "Database"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtDatabase 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmInitialPreferances"
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
    Dim FSys As New Scripting.FileSystemObject

Private Sub bttnExit_Click()
    On Error Resume Next
    End
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call SetColours(Me)
    Call SetPreferances
End Sub

Private Sub bttnClose_Click()
    On Error Resume Next
    Unload Me
    
End Sub

Private Sub SetPreferances()
    On Error Resume Next
    Dim TemResponce As Integer
    If FSys.FileExists(ProgramVariable.DatabaseName) = True Then
        txtDatabase.text = ProgramVariable.DatabaseName
    Else
        txtDatabase.text = "You have not selected a valid database"
        txtDatabase.ForeColor = vbYellow
        txtDatabase.BackColor = vbRed
    End If
End Sub


Private Sub SavePreferancesToFile()
    On Error Resume Next
    SaveSetting App.EXEName, "Options", "Database", txtDatabase.text
End Sub

Private Sub SavePreferancesToMemory()
    On Error Resume Next
    ProgramVariable.DatabaseName = txtDatabase.text
End Sub

Private Sub bttnSelectDatabasePath_Click()
    On Error Resume Next
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
    CommonDialog1.DefaultExt = "mdb"
    CommonDialog1.Filter = "MoH Database|MoLA.mdb"
    On Error Resume Next
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError = False Then
        txtDatabase.text = CommonDialog1.FileName
        SaveSetting App.EXEName, "Options", "Database", txtDatabase.text
        Unload Me
    Else
        MsgBox "You have not selected valid database. The program may not function", vbCritical, "No database"
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Dim TemResponce As Integer
    If FSys.FileExists(txtDatabase.text) = False Then
        MsgBox "You have not selected a valid database", vbCritical, "Database?"
        Cancel = True
        txtDatabase.SetFocus
        On Error Resume Next: SendKeys "{home}+{end}"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call SavePreferancesToFile
    Call SavePreferancesToMemory
End Sub
