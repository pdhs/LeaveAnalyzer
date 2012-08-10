Attribute VB_Name = "modSettings"
' Auther : Dr. M. H. B. Ariyaratne
'          buddhika.ari@gmail.com
'          buddhika_ari@yahoo.com
'          +94 71 58 12399
'          GPL Licence

Option Explicit

Public Sub SaveCommonSettings(myForm As Form)
    Dim MyCtrl As Control
    Dim i As Integer
    On Error Resume Next
    For Each MyCtrl In myForm.Controls
    
        SaveSetting App.EXEName, myForm.Name & MyCtrl.Name, "Top", MyCtrl.Top
        SaveSetting App.EXEName, myForm.Name & MyCtrl.Name, "Left", MyCtrl.Left
        SaveSetting App.EXEName, myForm.Name & MyCtrl.Name, "Width", MyCtrl.Width
        SaveSetting App.EXEName, myForm.Name & MyCtrl.Name, "Height", MyCtrl.Height
    
        If TypeOf MyCtrl Is MSFlexGrid Then
            For i = 0 To MyCtrl.Cols - 1
                SaveSetting App.EXEName, myForm.Name & MyCtrl.Name, i, MyCtrl.ColWidth(i)
            Next
        ElseIf TypeOf MyCtrl Is ComboBox Then
            If InStr(MyCtrl.Tag, "SS") Then
                SaveSetting App.EXEName, myForm.Name, MyCtrl.Name, MyCtrl.text
            End If
        ElseIf TypeOf MyCtrl Is DataCombo Then
            If InStr(MyCtrl.Tag, "SS") Then
                SaveSetting App.EXEName, myForm.Name, MyCtrl.Name, MyCtrl.text
            End If
        End If
    Next
    SaveSetting App.EXEName, myForm.Name, "Top", myForm.Top
    SaveSetting App.EXEName, myForm.Name, "Left", myForm.Left
    SaveSetting App.EXEName, myForm.Name, "Width", myForm.Width
    SaveSetting App.EXEName, myForm.Name, "Height", myForm.Height
    SaveSetting App.EXEName, myForm.Name, "WindowState", myForm.WindowState
    
End Sub

Public Sub GetCommonSettings(myForm As Form)
    Dim MyCtrl As Control
    Dim i As Integer
    
    On Error Resume Next
    
    For Each MyCtrl In myForm.Controls
    
        MyCtrl.Height = GetSetting(App.EXEName, myForm.Name & MyCtrl.Name, "Height", MyCtrl.Height)
        MyCtrl.Width = GetSetting(App.EXEName, myForm.Name & MyCtrl.Name, "Width", MyCtrl.Width)
        MyCtrl.Top = GetSetting(App.EXEName, myForm.Name & MyCtrl.Name, "Top", MyCtrl.Top)
        MyCtrl.Left = GetSetting(App.EXEName, myForm.Left & MyCtrl.Name, "Left", MyCtrl.Left)
    
    
        If TypeOf MyCtrl Is MSFlexGrid Then
            For i = 0 To MyCtrl.Cols - 1
                MyCtrl.ColWidth(i) = GetSetting(App.EXEName, myForm.Name & MyCtrl.Name, i, MyCtrl.ColWidth(i))
                MyCtrl.AllowUserResizing = flexResizeColumns
            Next
        ElseIf TypeOf MyCtrl Is ComboBox Then
            On Error Resume Next
            If InStr(MyCtrl.Tag, "SS") Then
                MyCtrl.text = GetSetting(App.EXEName, myForm.Name, MyCtrl.Name, "")
            End If
            On Error GoTo 0
        ElseIf TypeOf MyCtrl Is DataCombo Then
            On Error Resume Next
            If InStr(MyCtrl.Tag, "SS") Then
                MyCtrl.text = GetSetting(App.EXEName, myForm.Name, MyCtrl.Name, "")
            End If
            On Error GoTo 0
        End If

    Next
    
'    If Val(GetSetting(App.EXEName, myForm.Name, "Width", myForm.Top)) < myForm.Height * 0.75 Then myForm.Top = GetSetting(App.EXEName, myForm.Name, "Top", myForm.Top)
'    If Val(GetSetting(App.EXEName, myForm.Name, "Width", myForm.Left)) < myForm.Width * 0.75 Then myForm.Left = GetSetting(App.EXEName, myForm.Name, "Left", myForm.Left)
'    If Val(GetSetting(App.EXEName, myForm.Name, "Width", myForm.Width)) > 0 Then myForm.Width = GetSetting(App.EXEName, myForm.Name, "Width", myForm.Width)
'    If Val(GetSetting(App.EXEName, myForm.Name, "Width", myForm.Height)) > 0 Then myForm.Height = GetSetting(App.EXEName, myForm.Name, "Height", myForm.Height)

    myForm.Top = GetSetting(App.EXEName, myForm.Name, "Top", myForm.Top)
    myForm.Left = GetSetting(App.EXEName, myForm.Name, "Left", myForm.Left)
    myForm.Width = GetSetting(App.EXEName, myForm.Name, "Width", myForm.Width)
    myForm.Height = GetSetting(App.EXEName, myForm.Name, "Height", myForm.Height)


    On Error Resume Next
    myForm.WindowState = GetSetting(App.EXEName, myForm.Name, "WindowState", myForm.WindowState)

End Sub
