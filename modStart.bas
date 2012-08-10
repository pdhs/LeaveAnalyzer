Attribute VB_Name = "modStart"
' Auther : Dr. M. H. B. Ariyaratne
'          buddhika.ari@gmail.com
'          buddhika_ari@yahoo.com
'          +94 71 58 12399
'          GPL Licence

Option Explicit
    Dim FSys As New Scripting.FileSystemObject

Public Sub Main()
    Call setSecKey
    Call LoadColourPreferances
    Call LoadPreferances
    Call selectDatabase
    Call updateDatabase
    frmLogin.Show
End Sub

Public Sub updateDatabase()
    Dim dbVersion As New clsVersion
    dbVersion.loadData
    If dbVersion.VersionNo = 1 Then
        updateDBFromV1ToV2
    ElseIf dbVersion.VersionNo = 2 Then
    
    End If
End Sub

Private Sub updateDBFromV1ToV2()

End Sub

Private Sub setSecKey()
    Dim mySec As New clsSecurity
    ProgramVariable.SecurityKey = mySec.Decode(")?C{.‡", "Buddhika")
End Sub

Private Sub selectDatabase()
    If FSys.FileExists(ProgramVariable.DatabaseName) = True Then
    Dim cnn As New ADODB.Connection
        Dim constr As String
        constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & ProgramVariable.DatabaseName & " ;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
        ProgramVariable.conn.Open constr
    Else
        frmInitialPreferances.Show 1
        selectDatabase
    End If
End Sub

Private Sub LoadPreferances()
'On Error Resume Next
    ProgramVariable.DatabaseName = GetSetting(App.EXEName, "Options", "Database", App.Path & "\MoLA.mdb")

    ProgramVariable.LongDateFormat = GetSetting(App.EXEName, "Options", "LongDateFormat", "yyyy mmmm dd")
    ProgramVariable.ShortDateFormat = GetSetting(App.EXEName, "Options", "ShortDateFormat", "dd MM yy")

    If CBool(GetSetting(App.EXEName, "Options", "Energy", True)) = True Then
        DefaultColourScheme = Energy
    ElseIf CBool(GetSetting(App.EXEName, "Options", "Sunny", False)) = True Then
        DefaultColourScheme = Sunny
    ElseIf CBool(GetSetting(App.EXEName, "Options", "Aqua", False)) = True Then
        DefaultColourScheme = Aqua
    End If
        
    MDIImageFile = GetSetting(App.EXEName, "Options", "MDIImageFile", "")
    
End Sub


