VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVariables"
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
    Private mvarSecurityKey As String
    Private mvarDatabaseName As String 'local copy
    Private mvarDatabaseLocation As String 'local copy
    Private mvarconn As New ADODB.Connection 'local copy
    Private mvarShortDateFormat As String 'local copy
    Private mvarLongDateFormat As String 'local copy
    Private mvarloggedUser As clsUser 'local copy

Public Property Set loggedUser(ByVal vData As clsUser)
    Set mvarloggedUser = vData
End Property


Public Property Get loggedUser() As clsUser
    Set loggedUser = mvarloggedUser
End Property

Public Property Let LongDateFormat(ByVal vData As String)
    mvarLongDateFormat = vData
End Property


Public Property Get LongDateFormat() As String
    LongDateFormat = mvarLongDateFormat
End Property

Public Property Let ShortDateFormat(ByVal vData As String)
    mvarShortDateFormat = vData
End Property


Public Property Get ShortDateFormat() As String
    ShortDateFormat = mvarShortDateFormat
End Property

Public Property Set conn(ByVal vData As ADODB.Connection)
    Set mvarconn = vData
End Property


Public Property Get conn() As ADODB.Connection
    Set conn = mvarconn
End Property

Public Property Let DatabaseLocation(ByVal vData As String)
    mvarDatabaseLocation = vData
End Property

Public Property Get DatabaseLocation() As String
    DatabaseLocation = mvarDatabaseLocation
End Property

Public Property Let DatabaseName(ByVal vData As String)
    mvarDatabaseName = vData
End Property

Public Property Get DatabaseName() As String
    DatabaseName = mvarDatabaseName
End Property

Public Property Let SecurityKey(ByVal vData As String)
    mvarSecurityKey = vData
End Property


Public Property Get SecurityKey() As String
    SecurityKey = mvarSecurityKey
End Property



