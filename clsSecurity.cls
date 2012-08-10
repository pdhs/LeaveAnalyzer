VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Auther : Dr. M. H. B. Ariyaratne
'          buddhika.ari@gmail.com
'          buddhika_ari@yahoo.com
'          +94 71 58 12399
'          GPL Licence

Option Explicit

Public Function Hash(ByVal text As String) As String
    Dim a As Long
    Dim i As Long
    a = 1
    For i = 1 To Len(text)
        a = Sqr(a * i * Asc(Mid(text, i, 1)))
    Next i
    Rnd (-1)
    Randomize a
    For i = 1 To 16
        Hash = Hash & Chr(Int(Rnd * 256))
    Next i
End Function

Public Function Transform(ByVal text As String, _
    ByVal key As String) As String
    Dim i As Long
    Dim a As Long
    For i = 1 To Len(text)
      a = i Mod Len(key): If a = 0 Then a = Len(key)
      Transform = Transform & _
         Chr(Asc(Mid(key, a, 1)) Xor Asc(Mid(text, i, 1)))
    Next i
End Function


Public Function Encode(InputText As String, EKey As String) As String
    
    InputText = Transform(InputText, EKey)
    
    Dim User_Text                   As String
    Dim User_Text_Len               As Double
    
    Dim I_For                       As Double
    Dim Char_Val                    As String
    
    Dim Encode_Char                As String
    Dim Decode_Char                As String

    Dim Check_Val As Byte
    Check_Val = 0
    User_Text = ""
    User_Text = InputText
    
    Encode_Char = ""
    If User_Text = "" Then Exit Function
    User_Text_Len = Len(User_Text)
    
    For I_For = 1 To User_Text_Len
        Char_Val = Asc(Mid(User_Text, I_For, 1))
        DoEvents
        '--------------------------------------------
            Check_Val = 0: Check_Val = Oct(Char_Val)
        '--------------------------------------------
        If Check_Val >= 245 Then
            DoEvents
            Encode_Char = Encode_Char & Chr(Check_Val)
        Else
            DoEvents
            Check_Val = Check_Val + 10
            Encode_Char = Encode_Char & Chr(Check_Val)
        End If
    Next I_For
    
    Encode = ""
    Encode = StrReverse(Encode_Char)
    
End Function

Public Function Decode(InputText As String, DKey As String) As String
    Dim User_Text                   As String
    Dim User_Text_Len               As Double
    
    Dim I_For                       As Double
    Dim Char_Val                    As String
    
    Dim Encode_Char                As String
    Dim Decode_Char                As String
    Dim I_Oct                       As Double
    Dim Len_Char_Val                As Integer
    Dim Oct_Val                     As Integer
    Dim B_Oct                       As Integer
    Dim Temp                        As Byte
    
    User_Text = ""
    User_Text = StrReverse(InputText)
    Decode_Char = ""
    If User_Text = "" Then Exit Function

    User_Text_Len = Len(User_Text)
    
    For I_For = 1 To User_Text_Len
        Char_Val = Asc(Mid(User_Text, I_For, 1))
        '----------------------------------------------------
            If Char_Val <= 245 Then Char_Val = Char_Val - 10
            Len_Char_Val = Len(Char_Val)
        '----------------------------------------------------
        Oct_Val = 0
        For I_Oct = 0 To Len_Char_Val - 1
            DoEvents
            B_Oct = 8 ^ I_Oct
            If I_Oct = 0 Then Temp = B_Oct Else Temp = I_Oct + 1
            Oct_Val = CInt(Oct_Val) + CInt((B_Oct) * (Mid(StrReverse(Char_Val), Temp, 1)))
        Next I_Oct
        Decode_Char = Decode_Char & Chr(Oct_Val)
    Next I_For
    
    Decode = ""
    Decode = Transform(Decode_Char, DKey)
    

End Function




