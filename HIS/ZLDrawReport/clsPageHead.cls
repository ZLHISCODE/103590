VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPageHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'######################################################################################################################
Private mclsCommonLib As clsCommonLib

'######################################################################################################################

Public Function InsertText(ByVal strText As String, ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************

    mclsCommonLib.Rect = objRect
    mclsCommonLib.Font = objFont

    InsertText = AppendPrintData("ҳü", "�ı�", objFont.HAlignment, objFont.WordWarp, strText, objFont.VAlignment, objFont.MuliLine, objFont.Rows, objFont.AutoFit)
    
    objRect = mclsCommonLib.Rect
    objFont = mclsCommonLib.Font
    
End Function

Public Function InsertLine(ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mclsCommonLib.Rect = objRect
    mclsCommonLib.Font = objFont
    
    InsertLine = AppendPrintData("ҳü", "����")
    
    objRect = mclsCommonLib.Rect
    objFont = mclsCommonLib.Font

End Function

Public Function InsertPageNo(ByVal strText As String, ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mclsCommonLib.Rect = objRect
    mclsCommonLib.Font = objFont
    
    InsertPageNo = AppendPrintData("ҳ��", "ҳ��", objFont.HAlignment, False, strText, 2, False, 1, False)
    
    objRect = mclsCommonLib.Rect
    objFont = mclsCommonLib.Font
End Function

Public Function InsertGrid(ByVal objclsGrid As clsGrid, ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    mclsCommonLib.Rect = objRect
    mclsCommonLib.Font = objFont
    
    InsertGrid = objclsGrid.FormatData("ҳü")

    objRect = mclsCommonLib.Rect
    objFont = mclsCommonLib.Font
End Function

Private Sub Class_Initialize()
    Set mclsCommonLib = New clsCommonLib
End Sub

Private Sub Class_Terminate()
    Set mclsCommonLib = Nothing
End Sub