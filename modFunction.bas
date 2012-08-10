Attribute VB_Name = "modFunction"
' Auther : Dr. M. H. B. Ariyaratne
'          buddhika.ari@gmail.com
'          buddhika_ari@yahoo.com
'          +94 71 58 12399
'          GPL Licence

Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    Dim i As Integer
    Dim temSQL As String
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

    


Public Function RepeatString(InputString As String, RepeatNo As Integer) As String
    Dim r As Integer
    For r = 1 To RepeatNo
        RepeatString = RepeatString & InputString
    Next r
End Function




Public Sub GridToExcel(ExportGrid As MSFlexGrid, Optional Topic As String, Optional Subtopic As String)
    If ExportGrid.Rows <= 1 Then
        MsgBox "Noting to Export"
        Exit Sub
    End If
    
    Dim AppExcel As Excel.Application
    Dim myworkbook As Excel.Workbook
    Dim myWorkSheet1 As Excel.Worksheet
    Dim temRow As Integer
    Dim temCol As Integer
    
    Set AppExcel = CreateObject("Excel.Application")
    Set myworkbook = AppExcel.Workbooks.Add
    Set myWorkSheet1 = AppExcel.Worksheets(1)
    
    myWorkSheet1.Cells(1, 1) = Topic
    myWorkSheet1.Cells(2, 1) = Subtopic
    
    For temRow = 0 To ExportGrid.Rows - 1
        For temCol = 0 To ExportGrid.Cols - 1
            myWorkSheet1.Cells(temRow + 4, temCol + 1) = ExportGrid.TextMatrix(temRow, temCol)
        Next
    Next temRow
    
    myWorkSheet1.Range("A1:" & GetColumnName(CDbl(temCol)) & temRow + 2).AutoFormat Format:=xlRangeAutoFormatClassic1
    
    myWorkSheet1.Range("A" & temRow + 3 & ":" & GetColumnName(CDbl(temCol)) & temRow + 3).AutoFormat Format:=xlRangeAutoFormat3DEffects1
    
    Topic = "Day End Summery " & Format(Date, "dd MMMM yyyy")
    myworkbook.SaveAs (App.Path & "\" & Topic & ".xls")
    myworkbook.Save
    myworkbook.Close
    
    ShellExecute 0&, "open", App.Path & "\" & Topic & ".xls", "", "", vbMaximizedFocus
End Sub

Private Function GetColumnName(ColumnNo As Long) As String
    Dim temnum As Integer
    Dim temnum1 As Integer
    
    If ColumnNo < 27 Then
        GetColumnName = Chr(ColumnNo + 64)
    Else
        temnum = ColumnNo \ 26
        temnum1 = ColumnNo Mod 26
        GetColumnName = Chr(temnum + 64) & Chr(temnum1 + 64)
    End If
End Function



Public Function FillAnyGrid(InputSql As String, InputGrid As MSFlexGrid, TotalNameCol As Integer, TotalCols() As Integer, OmitRepeatCols() As Integer, Optional AddBlankLine As Boolean, Optional ColToAddBlankLineWhenNew As Integer, Optional AddLineInsteadOfBlank As Boolean, Optional SubtotalColForBlankLine As Integer) As Double()
    On Error GoTo eh
    
    Dim rsTem As New ADODB.Recordset
    Dim colTotal() As Double
    Dim previousValue() As String
    Dim previousValue2 As String
    Dim AddBlankThisTime As Boolean
    
    Dim i As Integer
    Dim col As Integer
    Dim noRepeat As Boolean
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = InputSql
        .Open temSQL, ProgramVariable.conn, adOpenStatic, adLockReadOnly
        
        InputGrid.Clear
        
        InputGrid.Rows = 1
        InputGrid.Cols = .Fields.Count
        
        ReDim colTotal(.Fields.Count)
        ReDim previousValue(.Fields.Count)
        
        InputGrid.row = 0
                    
        For i = 0 To .Fields.Count - 1
            InputGrid.col = i
            InputGrid.text = .Fields(i).Name
        Next i
        
        While .EOF = False
            InputGrid.Rows = InputGrid.Rows + 1
            InputGrid.row = InputGrid.Rows - 1
                        
            AddBlankThisTime = False
                        
            For i = 0 To .Fields.Count - 1
                InputGrid.col = i
                
                
                If i = ColToAddBlankLineWhenNew And AddBlankLine = True Then
                    If .AbsolutePosition = 1 Then
                        previousValue2 = .Fields(i).Value
                    End If
                    If previousValue2 <> .Fields(i).Value Then
                        InputGrid.Rows = InputGrid.Rows + 2
                        InputGrid.row = InputGrid.Rows - 1
                        previousValue2 = .Fields(i).Value
                    End If
                End If
                
                If UBound(OmitRepeatCols) > 0 Then
                    noRepeat = True
                    For col = 0 To UBound(OmitRepeatCols) - 1
                        If OmitRepeatCols(col) = i Then
                            noRepeat = False
                        End If
                    Next
                    If noRepeat = False Then
                        For col = 0 To UBound(OmitRepeatCols) - 1
                            If OmitRepeatCols(col) = i Then
                                If previousValue(i) <> .Fields(i).Value Then
                                    previousValue(i) = .Fields(i).Value
                                    If IsNull(.Fields(i).Value) = False Then
                                        InputGrid.text = .Fields(i).Value
                                    End If
                                
                                Else
                                    
                                End If
                                
                            Else
                            
                            End If
                        Next
                    Else
                        InputGrid.text = .Fields(i).Value
                    End If
                Else
                    If IsNull(.Fields(i).Value) = False Then
                        InputGrid.text = .Fields(i).Value
                    End If
                End If
                
                For col = 0 To UBound(TotalCols) - 1
                    If TotalCols(col) = i Then
                        If IsNull(.Fields(i).Value) = False Then
                            colTotal(i) = colTotal(i) + Val(.Fields(i).Value)
                        End If
                    End If
                Next
            
            Next i
            .MoveNext
        Wend
        .Close
    End With
    
    If UBound(TotalCols) > 0 Then
        InputGrid.Rows = InputGrid.Rows + 2
        InputGrid.row = InputGrid.Rows - 1
        InputGrid.col = TotalNameCol
        InputGrid.text = "Total"
        For i = 0 To InputGrid.Cols - 1
            InputGrid.col = i
            For col = 0 To UBound(TotalCols) - 1
                If TotalCols(col) = i Then
                    InputGrid.text = colTotal(i)
                End If
            Next
        Next i
    End If
    
    Dim temCol As Integer
    Dim temRow As Integer
    Dim temColTextLength() As Integer
    Dim SubTotal As Double
    Dim AllColsOfTheRowIsBlank As Boolean
    Dim temBlankColCount As Integer
    
    ReDim temColTextLength(InputGrid.Cols - 1)
    
    
    If AddLineInsteadOfBlank = True Then
        
        With InputGrid
            
            For temRow = 1 To .Rows - 1
                
                AllColsOfTheRowIsBlank = True
                
                For temCol = 0 To .Cols - 1
                    If Trim(.TextMatrix(temRow, temCol)) <> "" And temCol <> SubtotalColForBlankLine Then
                        AllColsOfTheRowIsBlank = False
                        
                    End If
                    If temCol = SubtotalColForBlankLine Then
                            SubTotal = SubTotal + Val(.TextMatrix(temRow, temCol))
                    End If
                Next temCol
                
                If AllColsOfTheRowIsBlank = True Then
                    temBlankColCount = temBlankColCount + 1
                End If
                
                
                If temBlankColCount = 2 Then
                    For temCol = 0 To .Cols - 1
                        .TextMatrix(temRow, temCol) = RepeatString("-", temColTextLength(temCol))
                    Next temCol
                    temBlankColCount = 0
                End If
                
                If temBlankColCount = 1 Then
                    For temCol = 0 To .Cols - 1
                        If temCol = SubtotalColForBlankLine Then
                            .TextMatrix(temRow, temCol) = SubTotal
                            SubTotal = 0
                        End If
                    Next temCol
                End If
                
                If AllColsOfTheRowIsBlank = False Then
                    For temCol = 0 To .Cols - 1
                        temColTextLength(temCol) = Len(.TextMatrix(temRow, temCol))
                    Next temCol
                End If
                
            Next temRow
        End With
    
    End If
    
    Dim temDbl As Double
    FillAnyGrid = colTotal

    Exit Function
    
eh:     MsgBox "Error" & vbNewLine & Err.Description

End Function
    
    
