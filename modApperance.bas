Attribute VB_Name = "modApperance"
' Auther : Dr. M. H. B. Ariyaratne
'          buddhika.ari@gmail.com
'          buddhika_ari@yahoo.com
'          +94 71 58 12399
'          GPL Licence


Option Explicit
    Public MDIImageFile As String
    
    Public Type ColourScheme
        LabelForeColour As Long
        LabelBackColour As Long
        ButtonForeColour As Long
        ButtonBackColour As Long
        ButtonBorderColour As Long
        GridLightBackColour As Long
        GridDarkBackColour As Long
        ComboBackColour As Long
        ComboForeColour As Long
        TextBackColour As Long
        TextForeColour As Long
    End Type
    
    Public Energy As ColourScheme
    Public Aqua As ColourScheme
    Public Sunny As ColourScheme

    Public DefaultColourScheme As ColourScheme
    
    
Public Sub LoadColourPreferances()
    On Error Resume Next
    Energy.LabelForeColour = RGB(154, 14, 20)
    Energy.LabelBackColour = RGB(255, 220, 168)
    Energy.ButtonBackColour = RGB(255, 129, 81)
    Energy.ButtonForeColour = RGB(154, 14, 20)
    Energy.ButtonBorderColour = RGB(154, 14, 20)
    Energy.GridDarkBackColour = RGB(255, 220, 168)
    Energy.GridLightBackColour = RGB(255, 245, 218)
    Energy.ComboForeColour = RGB(154, 14, 20)
    Energy.ComboBackColour = RGB(255, 245, 218)
    Energy.TextForeColour = RGB(154, 14, 20)
    Energy.TextBackColour = RGB(255, 245, 218)
    

    Sunny.LabelForeColour = RGB(154, 97, 14)
    Sunny.LabelBackColour = RGB(255, 236, 168)
    Sunny.ButtonBackColour = RGB(255, 200, 0)
    Sunny.ButtonForeColour = RGB(154, 97, 14)
    Sunny.ButtonBorderColour = RGB(154, 97, 14)
    Sunny.GridDarkBackColour = RGB(255, 236, 168)
    Sunny.GridLightBackColour = RGB(255, 255, 232)
    Sunny.ComboForeColour = RGB(154, 97, 14)
    Sunny.ComboBackColour = RGB(255, 255, 232)
    Sunny.TextForeColour = RGB(154, 97, 14)
    Sunny.TextBackColour = RGB(255, 255, 232)
    
    Aqua.LabelForeColour = RGB(34, 134, 84)
    Aqua.LabelBackColour = RGB(168, 212, 255)
    Aqua.ButtonBackColour = RGB(100, 225, 225)
    Aqua.ButtonForeColour = RGB(34, 134, 84)
    Aqua.ButtonBorderColour = RGB(34, 134, 84)
    Aqua.GridDarkBackColour = RGB(168, 212, 255)
    Aqua.GridLightBackColour = RGB(232, 255, 255)
    Aqua.ComboForeColour = RGB(34, 134, 84)
    Aqua.ComboBackColour = RGB(232, 255, 255)
    Aqua.TextForeColour = RGB(34, 134, 84)
    Aqua.TextBackColour = RGB(232, 255, 255)


End Sub
    
Public Sub SetColours(myForm As Form)
    On Error Resume Next
    myForm.ForeColor = DefaultColourScheme.LabelForeColour
    myForm.BackColor = DefaultColourScheme.LabelBackColour
    On Error Resume Next
    Dim MyControl As Control
    For Each MyControl In myForm.Controls
        If TypeOf MyControl Is ButtonEx Then
            MyControl.ForeColor = DefaultColourScheme.ButtonForeColour
            MyControl.BackColor = DefaultColourScheme.ButtonBackColour
            MyControl.BorderColor = DefaultColourScheme.ButtonBorderColour
        ElseIf TypeOf MyControl Is CommandButton Then
            MyControl.ForeColor = DefaultColourScheme.ButtonForeColour
            MyControl.BackColor = DefaultColourScheme.ButtonBackColour
            MyControl.BorderColor = DefaultColourScheme.ButtonBorderColour
        ElseIf TypeOf MyControl Is ListBox Or TypeOf MyControl Is DataList Then
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
        ElseIf TypeOf MyControl Is Label Then
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
        ElseIf TypeOf MyControl Is DataCombo Or TypeOf MyControl Is ComboBox Then
            MyControl.ForeColor = DefaultColourScheme.ComboForeColour
            MyControl.BackColor = DefaultColourScheme.ComboBackColour
        ElseIf TypeOf MyControl Is TextBox Then
            MyControl.ForeColor = DefaultColourScheme.TextForeColour
            MyControl.BackColor = DefaultColourScheme.TextBackColour
        ElseIf TypeOf MyControl Is DTPicker Then
            MyControl.ForeColor = DefaultColourScheme.TextForeColour
            MyControl.BackColor = DefaultColourScheme.TextBackColour
        ElseIf TypeOf MyControl Is MSFlexGrid Then
            ColourGrid MyControl
        Else
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
            MyControl.BackStyle = 0
        End If
    Next
End Sub


Public Sub ColourGrid(GridToColour As MSFlexGrid)
    Dim col As Integer
    Dim row As Integer
    With GridToColour
        .Visible = False
        For row = 0 To .Rows - 1
            For col = 0 To .Cols - 1
                .col = col
                .row = row
                If .row Mod 2 = 0 Then
                    .CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    .CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
            Next
        Next
        .Visible = True
    End With
End Sub

Public Sub prepareEdit(editControlls As Collection, selectControlls As Collection)
    enableControls editControlls, True
    enableControls selectControlls, False
End Sub

Private Sub enableControls(myControlls As Collection, toEnable As Boolean)
    Dim temControl As Control
    On Error Resume Next
    For Each temControl In myControlls
        temControl.Enabled = toEnable
    Next
End Sub


Public Sub prepareSelect(editControlls As Collection, selectControlls As Collection)
    enableControls editControlls, False
    enableControls selectControlls, True
End Sub

Public Sub clearValues(myControlls As Collection)
    Dim temControl As Control
    On Error Resume Next
    For Each temControl In myControlls
        temControl.text = Empty
    Next
End Sub
