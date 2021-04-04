VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisReportFoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
Private mclsLisReportLib As clsLisReportLib

'######################################################################################################################

Public Function InsertText(ByVal strText As String, ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mclsLisReportLib.RECT = objRect
    mclsLisReportLib.Font = objFont

    InsertText = AppendPrintData("页脚", "文本", objFont.HAlignment, objFont.WordWarp, strText, objFont.VAlignment, objFont.MuliLine, objFont.Rows, objFont.AutoFit)
    
    objRect = mclsLisReportLib.RECT
    objFont = mclsLisReportLib.Font
    
End Function

Public Function InsertLine(ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mclsLisReportLib.RECT = objRect
    mclsLisReportLib.Font = objFont
    
    InsertLine = AppendPrintData("页脚", "线条")
    
    objRect = mclsLisReportLib.RECT
    objFont = mclsLisReportLib.Font
    
End Function

Public Function InsertGrID(ByVal objclsGrID As clsGrid, ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mclsLisReportLib.RECT = objRect
    mclsLisReportLib.Font = objFont
    
    InsertGrID = objclsGrID.FormatData("页脚")

    objRect = mclsLisReportLib.RECT
    objFont = mclsLisReportLib.Font
End Function

Public Function InsertPageNo(ByVal strText As String, ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mclsLisReportLib.RECT = objRect
    mclsLisReportLib.Font = objFont
    
    InsertPageNo = AppendPrintData("页脚", "页码", objFont.HAlignment, False, strText, 2, False, 1, False)
    
    objRect = mclsLisReportLib.RECT
    objFont = mclsLisReportLib.Font
End Function

Private Sub Class_Initialize()
    Set mclsLisReportLib = New clsLisReportLib
End Sub

Private Sub Class_Terminate()
    Set mclsLisReportLib = Nothing
End Sub
