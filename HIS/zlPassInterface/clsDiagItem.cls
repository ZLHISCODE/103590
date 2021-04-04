VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiagItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mstr¼²²¡±àÂë As String
Private mstrÕï¶Ï±àÂë As String
Private mstrÕï¶ÏÃèÊö As String
Private mstrÕï¶ÏID As String
Private mstr¼²²¡ID As String
Private mstrÕï¶ÏÀàĞÍ As String
Private mstrÕï¶ÏÊ±¼ä As String

Public Property Let str¼²²¡±àÂë(ByVal strData As String)
    mstr¼²²¡±àÂë = strData
End Property

Public Property Get str¼²²¡±àÂë() As String
    str¼²²¡±àÂë = mstr¼²²¡±àÂë
End Property

Public Property Let strÕï¶Ï±àÂë(ByVal strData As String)
    mstrÕï¶Ï±àÂë = strData
End Property

Public Property Get strÕï¶Ï±àÂë() As String
    strÕï¶Ï±àÂë = mstrÕï¶Ï±àÂë
End Property

Public Property Let strÕï¶ÏÃèÊö(ByVal strData As String)
    mstrÕï¶ÏÃèÊö = strData
End Property

Public Property Get strÕï¶ÏÃèÊö() As String
    strÕï¶ÏÃèÊö = mstrÕï¶ÏÃèÊö
End Property

Public Property Let str¼²²¡ID(ByVal strData As String)
    mstr¼²²¡ID = strData
End Property

Public Property Get str¼²²¡ID() As String
    str¼²²¡ID = mstr¼²²¡ID
End Property

Public Property Let strÕï¶ÏID(ByVal strData As String)
    mstrÕï¶ÏID = strData
End Property

Public Property Get strÕï¶ÏID() As String
    strÕï¶ÏID = mstrÕï¶ÏID
End Property

Public Property Let strÕï¶ÏÀàĞÍ(ByVal strData As String)
    mstrÕï¶ÏÀàĞÍ = strData
End Property

Public Property Get strÕï¶ÏÀàĞÍ() As String
    strÕï¶ÏÀàĞÍ = mstrÕï¶ÏÀàĞÍ
End Property

Public Property Let strÕï¶ÏÊ±¼ä(ByVal strData As String)
    mstrÕï¶ÏÊ±¼ä = strData
End Property

Public Property Get strÕï¶ÏÊ±¼ä() As String
    strÕï¶ÏÊ±¼ä = mstrÕï¶ÏÊ±¼ä
End Property
