VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResizer"
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

' if True, Control's Fontsizes are Changed (unless AllowFontChange is NO '
Private bFontResize As Boolean

' if True, form's height/width ratio is preserved '
Private bKeepRatio As Boolean

Private Type TControlInfo
       
       ctrl As Control
       Left As Single
       Top As Single
       Width As Single
       Height As Single
       FontSize As Single
       
End Type

Private Type TAllowChanges
  
       AllowChangeTop As Boolean
       AllowChangeLeft As Boolean
       AllowChangeWidth As Boolean
       AllowChangeHeight As Boolean
       AllowFontChange As Boolean
       
End Type

' this array holds the original position  '
' and size of all controls on parent form '
Dim Controls() As TControlInfo

' a reference to the parent form '
'Private WithEvents ParentForm As Form

' parent form's size at load time '
Private ParentWidth As Single
Private ParentHeight As Single

' ratio of original height/width '
Private HeightWidthRatio As Single
Property Let FontResize(ByVal DoFontResize As Boolean)
   
  bFontResize = DoFontResize

End Property
Public Sub InitializeResizer(frm As Form)

  ' the ParentWidth variable works as a flag '
  ParentWidth = 0
  
  ' save original ratio '
  HeightWidthRatio = frm.Height / frm.Width
  
End Sub


Public Sub FormResized(frm As Form)
  
  If ParentWidth = 0 Then
    
    Call Rebuild(frm)
  
  Else
    
    Call Refresh(frm)
  
  End If
  
End Sub
Private Function CheckForChanges(ByVal TagToUse As String) As TAllowChanges
  
  Dim ChangesToAllow As TAllowChanges
    
  On Error Resume Next
  
  ChangesToAllow.AllowChangeTop = True
  ChangesToAllow.AllowChangeLeft = True
  ChangesToAllow.AllowChangeWidth = True
  ChangesToAllow.AllowChangeHeight = True
  ChangesToAllow.AllowFontChange = IIf(bFontResize = True, True, False)
  
  If TagToUse <> "" Then
    
    If UCase(Left(TagToUse, 9)) = "MSIRESIZE" Then
      
      ChangesToAllow.AllowChangeTop = False
      ChangesToAllow.AllowChangeLeft = False
      ChangesToAllow.AllowChangeWidth = False
      ChangesToAllow.AllowChangeHeight = False
      ChangesToAllow.AllowFontChange = False
    
      If Mid(TagToUse, 10, 1) = "Y" Then
      
        ChangesToAllow.AllowChangeLeft = True
        
      End If
      
      If Mid(TagToUse, 11, 1) = "Y" Then
      
        ChangesToAllow.AllowChangeTop = True
        
      End If
      
      If Mid(TagToUse, 12, 1) = "Y" Then
      
        ChangesToAllow.AllowChangeWidth = True
        
      End If
      
      If Mid(TagToUse, 13, 1) = "Y" Then
      
        ChangesToAllow.AllowChangeHeight = True
        
      End If
      
      If Mid(TagToUse, 14, 1) = "Y" Then
      
        ChangesToAllow.AllowFontChange = True
        
      End If
      
    End If
    
  End If
  
  CheckForChanges = ChangesToAllow
  
End Function
Property Let KeepRatio(ByVal KeepControlRatio As Boolean)
  
  bKeepRatio = KeepControlRatio
  
End Property

Property Get KeepRatio() As Boolean
  
  KeepRatio = bKeepRatio
  
End Property
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' save size and position of all controls on parent form                  '
' you should manually invoke this method each time you add a new control '
' to the form (through Load method of a control array)                   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Rebuild(ByVal ParentForm As Form)
  
  ' rebuild the internal table '
  Dim i As Integer
  Dim ctrl As Control
  
  ' this is necessary for controls that don't support '
  ' all properties (e.g. Timer controls)              '
  On Error Resume Next
    
  ' save a reference to the parent form, and its initial size '
  ParentWidth = ParentForm.ScaleWidth
  ParentHeight = ParentForm.ScaleHeight
    
  ' read the position of all controls on the parent form '
  ReDim Controls(ParentForm.Controls.Count - 1) As TControlInfo
    
  For i = 0 To ParentForm.Controls.Count - 1
     
     Set ctrl = ParentForm.Controls(i)
        
     Set Controls(i).ctrl = ctrl
                     
     Controls(i).Left = ctrl.Left
     Controls(i).Top = ctrl.Top
     Controls(i).Width = ctrl.Width
     Controls(i).Height = ctrl.Height
     Controls(i).FontSize = ctrl.Font.Size
     
  Next
  
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' update size and position of controls on parent form '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Refresh(ByVal ParentForm As Form)
  
  Dim i As Integer
  Dim ctrl As Control
  Dim minFactor As Single
  Dim widthFactor As Single
  Dim heightFactor As Single
  
  Dim Changes As TAllowChanges
      
  ' this is necessary for controls that don't support '
  ' all properties (e.g. Timer controls)              '
  On Error Resume Next
  
  ' inhibits recursive calls if KeepRatio = True '
  Static Executing As Boolean
  
  If Executing Then
    
    Exit Sub
    
  End If
    
  If ParentForm.WindowState = vbMinimized Then
    
    Exit Sub
    
  End If
  
  If KeepRatio Then
    
    Executing = True
    
    ' we must keep original ratio '
    ParentForm.Height = HeightWidthRatio * ParentForm.Width
    Executing = False
  
  End If
    
  widthFactor = ParentForm.ScaleWidth / ParentWidth
  heightFactor = ParentForm.ScaleHeight / ParentHeight
  
  ' take the lesser of the two '
  If widthFactor < heightFactor Then
    
    minFactor = widthFactor
  
  Else
    
    minFactor = heightFactor
  
  End If
    
  ' this is a regular resize '
  For i = 0 To UBound(Controls)
        
     Changes = CheckForChanges(Controls(i).ctrl.Tag)
     
     ' the change of font must occur *before* the resizing '
     ' to account for companion scrollbar of listbox       '
     ' and other similar controls                          '
     If Changes.AllowFontChange = True Then
                       
       Controls(i).ctrl.Font.Size = Controls(i).FontSize * minFactor
                     
     End If
                     
     ' move and resize the controls - we can't use a Move       '
     ' method because some controls do not support the change   '
     ' of all the four properties (e.g. Height with comboboxes) '
     If Changes.AllowChangeLeft = True Then
                       
       Controls(i).ctrl.Left = Controls(i).Left * widthFactor
                     
     End If
                     
     If Changes.AllowChangeTop = True Then
                       
       Controls(i).ctrl.Top = Controls(i).Top * heightFactor
                     
     End If
                     
     If Changes.AllowChangeWidth = True Then
                       
       Controls(i).ctrl.Width = Controls(i).Width * widthFactor
                     
     End If
                     
     If Changes.AllowChangeHeight = True Then
                       
       Controls(i).ctrl.Height = Controls(i).Height * heightFactor
                     
     End If
     
  Next
  
End Sub
Property Get FontResize() As Boolean

   FontResize = bFontResize

End Property


