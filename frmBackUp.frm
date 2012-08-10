VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmBackUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6435
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBackUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   6435
   Begin VB.Frame Frame1 
      Caption         =   "Backup Directory"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtPath 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6015
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnBackUP 
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Back Up"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx bttnSelectDirectory 
      Height          =   495
      HelpContextID   =   10
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Select Directory"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    Private Const BIF_RETURNONLYFSDIRS = 1
    Private Const BIF_DONTGOBELOWDOMAIN = 2
    Private Const MAX_PATH = 260
    Private Declare Function SHBrowseForFolder Lib "shell32" _
                                      (lpbi As BrowseInfo) As Long
    
    Private Declare Function SHGetPathFromIDList Lib "shell32" _
                                      (ByVal pidList As Long, _
                                      ByVal lpBuffer As String) As Long
    
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
                                      (ByVal lpString1 As String, ByVal _
                                      lpString2 As String) As Long
    
    Private Type BrowseInfo
       hWndOwner      As Long
       pIDLRoot       As Long
       pszDisplayName As Long
       lpszTitle      As Long
       ulFlags        As Long
       lpfnCallback   As Long
       lparam         As Long
       iImage         As Long
    End Type


Private Sub bttnBackUP_Click()
    Dim TemResponce  As Integer
    On Error GoTo ErrorHandler
    If FSys.FolderExists(txtPath.text) = True Then
        Me.MousePointer = vbHourglass
        FSys.CopyFile ProgramVariable.DatabaseName, txtPath.text & "\LeMx BACKUP " & Format(Date, "dd mmm yy") & " " & Format(Time, "HH MM SS AMPM") & ".mdb", True
        TemResponce = MsgBox("Backup Successful", vbInformation, "Success")
        Me.MousePointer = vbDefault
    Else
        TemResponce = MsgBox("The path you selected is not valid. Please select a valid path", vbCritical, "Path not valid")
        Exit Sub
    End If
    Exit Sub
ErrorHandler:
        TemResponce = MsgBox("An unknown error occured. Please contact Fintec with following details." & vbNewLine & App.EXEName & vbNewLine & Me.Caption & vbNewLine & Err.Number & vbNewLine & Err.Description, vbInformation, Err.Description)
        Me.MousePointer = vbDefault
        Exit Sub
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnSelectDirectory_Click()
         Dim lpIDList As Long
         Dim sBuffer As String
         Dim szTitle As String
         Dim tBrowseInfo As BrowseInfo
         szTitle = "Select Backup Directory"
         With tBrowseInfo
            .hWndOwner = Me.hwnd
            .lpszTitle = lstrcat(szTitle, "")
            .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
         End With
         lpIDList = SHBrowseForFolder(tBrowseInfo)
         If (lpIDList) Then
            sBuffer = Space(MAX_PATH)
            SHGetPathFromIDList lpIDList, sBuffer
            sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            txtPath.text = sBuffer
         End If
End Sub

Private Sub Form_Load()
    Call SetColours(Me)
    GetCommonSettings Me
    
    txtPath.text = App.Path
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = vbRightButton Then
'        Me.PopupMenu MDIMain.mnuRxPopUp
'    End If
End Sub
