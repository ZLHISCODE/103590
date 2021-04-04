VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisReportData"
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

    InsertText = AppendPrintData("数据", "文本", objFont.HAlignment, objFont.WordWarp, strText, objFont.VAlignment, objFont.MuliLine, objFont.Rows, objFont.AutoFit)
    
    objRect = mclsLisReportLib.RECT
    objFont = mclsLisReportLib.Font
    
End Function

Public Function InsertContinueText(ByVal strText As String, ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mclsLisReportLib.RECT = objRect
    mclsLisReportLib.Font = objFont

    InsertContinueText = AppendPrintData("数据", "续页", objFont.HAlignment, objFont.WordWarp, strText, objFont.VAlignment, objFont.MuliLine, objFont.Rows, objFont.AutoFit)
    
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
    
    InsertLine = AppendPrintData("数据", "线条")
    
    objRect = mclsLisReportLib.RECT
    objFont = mclsLisReportLib.Font

End Function

Public Function InsertBackColor(ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mclsLisReportLib.RECT = objRect
    mclsLisReportLib.Font = objFont
    
    InsertBackColor = AppendPrintData("数据", "背景", 2, , , 2)
    
    objRect = mclsLisReportLib.RECT
    objFont = mclsLisReportLib.Font

End Function

Public Function InsertImage(ByRef objRect As USERRECT, ByRef objFont As USERFONT, ByVal strImageFile As String, Optional ByVal bytHAlignment As Byte = 2, Optional ByVal bytVAlignment As Byte = 2) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mclsLisReportLib.RECT = objRect
    mclsLisReportLib.Font = objFont
    
    InsertImage = AppendPrintData("数据", "图形", bytHAlignment, , strImageFile, bytVAlignment)
    
    objRect = mclsLisReportLib.RECT
    objFont = mclsLisReportLib.Font

End Function

Public Function InsertRectangle(ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objTmpRect As USERRECT
    
    mclsLisReportLib.RECT = objRect
    mclsLisReportLib.Font = objFont
    
    objTmpRect = objRect
    
    '左竖线
    objTmpRect.X0 = objRect.X0
    objTmpRect.Y0 = objRect.Y0
    objTmpRect.X1 = objRect.X0
    objTmpRect.Y1 = objRect.Y1
    mclsLisReportLib.RECT = objTmpRect
    InsertRectangle = AppendPrintData("数据", "线条")
    
    '上横线
    objTmpRect.X0 = objRect.X0
    objTmpRect.Y0 = objRect.Y0
    objTmpRect.X1 = objRect.X1
    objTmpRect.Y1 = objRect.Y0
    mclsLisReportLib.RECT = objTmpRect
    InsertRectangle = AppendPrintData("数据", "线条")
    
    '右竖线
    objTmpRect.X0 = objRect.X1
    objTmpRect.Y0 = objRect.Y0
    objTmpRect.X1 = objRect.X1
    objTmpRect.Y1 = objRect.Y1
    mclsLisReportLib.RECT = objTmpRect
    InsertRectangle = AppendPrintData("数据", "线条")
    
    '下横线
    objTmpRect.X0 = objRect.X0
    objTmpRect.Y0 = objRect.Y1
    objTmpRect.X1 = objRect.X1
    objTmpRect.Y1 = objRect.Y1
    mclsLisReportLib.RECT = objTmpRect
    InsertRectangle = AppendPrintData("数据", "线条")
    
    objRect = mclsLisReportLib.RECT
    objFont = mclsLisReportLib.Font

End Function

Public Function InsertGrID(ByVal objclsGrID As clsGrid, ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    objRect.X1 = objRect.X0 + objclsGrID.SumGrIDColWidth
    objRect.Y1 = objRect.Y0 + objclsGrID.SumGrIDRowHeight
    
    mclsLisReportLib.RECT = objRect
    mclsLisReportLib.Font = objFont
    
    InsertGrID = objclsGrID.FormatData("数据")
    
    objRect = mclsLisReportLib.RECT
    objFont = mclsLisReportLib.Font
    
End Function

Public Function NewCoverPage(Optional ByVal strShow As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim objRect As USERRECT
    
    On Error GoTo errHand

    objRect = mclsLisReportLib.RECT
    objRect.Page = objRect.Page + 1
    mclsLisReportLib.RECT = objRect
        
    NewCoverPage = InsertPage(objRect.Page, 0, strShow, False, False)
        
    Exit Function

errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function NewListPage(Optional ByVal strShow As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim objRect As USERRECT
    
    On Error GoTo errHand

    objRect = mclsLisReportLib.RECT
    objRect.Page = objRect.Page + 1
    mclsLisReportLib.RECT = objRect
        
    NewListPage = InsertPage(objRect.Page, 0, strShow, False, False)
        
    Exit Function

errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function NewPage(Optional ByVal bytCalc As Byte = 1, Optional ByVal strShow As String, Optional ByVal blnShowPageHead As Boolean = True, Optional ByVal blnShowPageFoot As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim objRect As USERRECT
    
    On Error GoTo errHand

    objRect = mclsLisReportLib.RECT
    objRect.Page = objRect.Page + 1
    mclsLisReportLib.RECT = objRect
        
    NewPage = InsertPage(objRect.Page, bytCalc, strShow, blnShowPageHead, blnShowPageFoot)
        
    Exit Function

errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If

End Function

Public Function AppendList(ByVal strListName As String, ByVal bytList As Byte, ByVal intPage As Integer) As Boolean
    '******************************************************************************************************************
    '功能：添加到目录索引
    '参数：
    '返回：
    '******************************************************************************************************************

    AppendList = AppendListData(strListName, bytList, intPage)
    
End Function

Public Property Let StartPage(ByVal intPage As Integer)
    '******************************************************************************************************************
    '功能：开始页码
    '参数：
    '返回：
    '******************************************************************************************************************
    
    gintStartPage = intPage
    
End Property

Public Function GetNewTempFile() As String
    
    Dim strFile As String
    
    strFile = "zlPublicLIS" & Format(grsTempFile.RecordCount, "000000") & ".tmp"
    
    strFile = CreateTmpFile(strFile)
    
    GetNewTempFile = strFile
    
    grsTempFile.AddNew
    grsTempFile("文件").Value = strFile
    
End Function

Private Sub Class_Initialize()
    Set mclsLisReportLib = New clsLisReportLib
End Sub

Private Sub Class_Terminate()
    Set mclsLisReportLib = Nothing
End Sub

