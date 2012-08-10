VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmRestore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restore"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRestore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6480
   Begin VB.Frame Frame1 
      Caption         =   "Select the directory from which to restore"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtPath 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   5895
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5106
      _Version        =   393216
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
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx bttnRestore 
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Restore"
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
   Begin btButtonEx.ButtonEx bttnSelectPath 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "Select& &Path"
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
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    Dim TemResponce  As Integer
    Dim ThisFolder As Folder
    Dim AllFiles As Files
    Dim ThisFile As File
    Dim TemString1 As String
    Dim TemString2 As String
    Dim TemString3 As String
    Dim TemString4 As String
    Dim TemDate As Date
    Private Const BIF_RETURNONLYFSDIRS = 1
    Private Const BIF_DONTGOBELOWDOMAIN = 2
    Private Const MAX_PATH = 260
    Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
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

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnRestore_Click()
    Grid1.col = 0
    TemResponce = MsgBox("Are you sure you want replace the current database with a previously backed-up database ", vbInformation + vbYesNo, "Restore?")
    If TemResponce = vbNo Then Exit Sub
    On Error GoTo ErrorHandler
    Me.MousePointer = vbHourglass
    DoEvents
    FSys.CopyFile ProgramVariable.DatabaseName, txtPath.text & "\LeMx BACKUP " & Format(Date, "dd mmm yy") & " " & Format(Time, "HH MM SS AMPM") & ".mdb", True
    FSys.CopyFile Grid1.text, ProgramVariable.DatabaseName, True
    
    TemResponce = MsgBox("Restore Successful", vbInformation, "Success")
    Me.MousePointer = vbDefault
    Exit Sub
ErrorHandler:
    TemResponce = MsgBox("An unknown error occured. Please contact Fintec with following details." & vbNewLine & App.EXEName & vbNewLine & Me.Caption & vbNewLine & Err.Number & vbNewLine & Err.Description, vbInformation, "Error")
    Exit Sub
End Sub

Private Sub ListDatabases()
    Dim NowROw As Long
    Grid1.Clear
    Grid1.Rows = 1
    Grid1.Cols = 3
    Grid1.ColWidth(0) = 1
    Grid1.ColWidth(1) = 3000
    Grid1.ColWidth(2) = Grid1.Width - 2101
    
    Grid1.row = 0
    
    Grid1.col = 1
    Grid1.text = "Last Modified Date"
    
    Grid1.col = 2
    Grid1.text = "Last Modified Time"
    
    If FSys.FolderExists(txtPath.text) = True Then
        
        Set ThisFolder = FSys.GetFolder(txtPath.text)
        Set AllFiles = ThisFolder.Files
        
        NowROw = 0
        For Each ThisFile In AllFiles
            If InStr(ThisFile.Name, "LeMx BACKUP ") > 0 Then
                NowROw = NowROw + 1
                Grid1.Rows = NowROw + 1
                Grid1.row = NowROw
                TemDate = ThisFile.DateLastModified
                Grid1.col = 0
                Grid1.text = ThisFile.Path
                If Grid1.row Mod 2 = 0 Then
                    Grid1.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    Grid1.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                Grid1.col = 1
                Grid1.CellAlignment = 4
                Grid1.text = Format(TemDate, ProgramVariable.LongDateFormat)
                If Grid1.row Mod 2 = 0 Then
                    Grid1.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    Grid1.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
            
                Grid1.col = 2
                Grid1.CellAlignment = 4
                Grid1.text = Format(TemDate, "HH MM SS AMPM")
                If Grid1.row Mod 2 = 0 Then
                    Grid1.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    Grid1.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
            
            
            End If
        Next
    End If
    bttnRestore.Enabled = False
End Sub

Private Sub bttnSelectPath_Click()
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
            Call ListDatabases
         End If
End Sub

Private Sub SetColours()
    Me.ForeColor = DefaultColourScheme.LabelForeColour
    Me.BackColor = DefaultColourScheme.LabelBackColour
    On Error Resume Next
    Dim MyControl As Control
    For Each MyControl In Controls
        If InStr(UCase(MyControl.Name), "BTN") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.ButtonForeColour
            MyControl.BackColor = DefaultColourScheme.ButtonBackColour
            MyControl.BorderColor = DefaultColourScheme.ButtonBorderColour
        ElseIf InStr(UCase(MyControl.Name), "LST") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
        ElseIf InStr(UCase(MyControl.Name), "TXTID") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
        ElseIf InStr(UCase(MyControl.Name), "CMB") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.ComboForeColour
            MyControl.BackColor = DefaultColourScheme.ComboBackColour
        ElseIf InStr(UCase(MyControl.Name), "TXT") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.TextForeColour
            MyControl.BackColor = DefaultColourScheme.TextBackColour
        ElseIf InStr(UCase(MyControl.Name), "DTP") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.TextForeColour
            MyControl.BackColor = DefaultColourScheme.TextBackColour
        Else
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
            MyControl.BackStyle = 0
        End If
    Next
End Sub

Private Sub Form_Load()
    GetCommonSettings Me
    
    Call SetColours
    bttnRestore.Enabled = False
    txtPath.text = GetSetting(App.EXEName, Me.Name, txtPath.Name, App.Path)
    Call ListDatabases
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    
    SaveSetting App.EXEName, Me.Name, txtPath.Name, txtPath.text
End Sub

Private Sub Grid1_Click()
    If Grid1.row < 1 Then
        bttnRestore.Enabled = False
        Exit Sub
    Else
        bttnRestore.Enabled = True
    End If
End Sub
