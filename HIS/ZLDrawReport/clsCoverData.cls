VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCoverData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
Private mclsCommonLib As clsCommonLib

'######################################################################################################################

Public Function InsertText(ByVal strText As String, ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    mclsCommonLib.Rect = objRect
    mclsCommonLib.Font = objFont

    InsertText = AppendPrintData("封面", "文本", objFont.HAlignment, objFont.WordWarp, strText, objFont.VAlignment, objFont.MuliLine, objFont.Rows, objFont.AutoFit)
    
    objRect = mclsCommonLib.Rect
    objFont = mclsCommonLib.Font
    
End Function

Public Function InsertLine(ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mclsCommonLib.Rect = objRect
    mclsCommonLib.Font = objFont
    
    InsertLine = AppendPrintData("封面", "线条")
    
    objRect = mclsCommonLib.Rect
    objFont = mclsCommonLib.Font

End Function

Public Function InsertBackColor(ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mclsCommonLib.Rect = objRect
    mclsCommonLib.Font = objFont
    
    InsertBackColor = AppendPrintData("封面", "背景", 2, , , 2)
    
    objRect = mclsCommonLib.Rect
    objFont = mclsCommonLib.Font

End Function

Public Function InsertImage(ByRef objRect As USERRECT, ByRef objFont As USERFONT, ByVal strImageFile As String, Optional ByVal bytHAlignment As Byte = 2, Optional ByVal bytVAlignment As Byte = 2) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mclsCommonLib.Rect = objRect
    mclsCommonLib.Font = objFont
    
    InsertImage = AppendPrintData("封面", "图形", bytHAlignment, , strImageFile, bytVAlignment)
    
    objRect = mclsCommonLib.Rect
    objFont = mclsCommonLib.Font

End Function

Public Function InsertRectangle(ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objTmpRect As USERRECT
    
    mclsCommonLib.Rect = objRect
    mclsCommonLib.Font = objFont
    
    objTmpRect = objRect
    
    '左竖线
    objTmpRect.X0 = objRect.X0
    objTmpRect.Y0 = objRect.Y0
    objTmpRect.X1 = objRect.X0
    objTmpRect.Y1 = objRect.Y1
    mclsCommonLib.Rect = objTmpRect
    InsertRectangle = AppendPrintData("封面", "线条")
    
    '上横线
    objTmpRect.X0 = objRect.X0
    objTmpRect.Y0 = objRect.Y0
    objTmpRect.X1 = objRect.X1
    objTmpRect.Y1 = objRect.Y0
    mclsCommonLib.Rect = objTmpRect
    InsertRectangle = AppendPrintData("封面", "线条")
    
    '右竖线
    objTmpRect.X0 = objRect.X1
    objTmpRect.Y0 = objRect.Y0
    objTmpRect.X1 = objRect.X1
    objTmpRect.Y1 = objRect.Y1
    mclsCommonLib.Rect = objTmpRect
    InsertRectangle = AppendPrintData("封面", "线条")
    
    '下横线
    objTmpRect.X0 = objRect.X0
    objTmpRect.Y0 = objRect.Y1
    objTmpRect.X1 = objRect.X1
    objTmpRect.Y1 = objRect.Y1
    mclsCommonLib.Rect = objTmpRect
    InsertRectangle = AppendPrintData("封面", "线条")
    
    objRect = mclsCommonLib.Rect
    objFont = mclsCommonLib.Font

End Function

Public Function InsertGrid(ByVal objclsGrid As clsGrid, ByRef objRect As USERRECT, ByRef objFont As USERFONT) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************

    objRect.X1 = objRect.X0 + objclsGrid.SumGridColWidth
    objRect.Y1 = objRect.Y0 + objclsGrid.SumGridRowHeight
    
    mclsCommonLib.Rect = objRect
    mclsCommonLib.Font = objFont
    
    InsertGrid = objclsGrid.FormatData("封面")
    
    objRect = mclsCommonLib.Rect
    objFont = mclsCommonLib.Font
    
End Function

Public Function InsertCurve(ByVal objclsCurve As clsCurve, Optional ByVal strPrex As String = "A") As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    InsertCurve = objclsCurve.FormatData("封面", strPrex)
    
End Function

Public Function NewCoverPage(Optional ByVal strShow As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim objRect As USERRECT
    
    On Error GoTo errHand

    objRect = mclsCommonLib.Rect
    objRect.Page = objRect.Page + 1
    mclsCommonLib.Rect = objRect
        
    NewCoverPage = InsertPage(objRect.Page, 0, strShow, False, False)
        
    Exit Function

errHand:
    If ErrCenter = 1 Then
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

    objRect = mclsCommonLib.Rect
    objRect.Page = objRect.Page + 1
    mclsCommonLib.Rect = objRect
        
    NewListPage = InsertPage(objRect.Page, 0, strShow, False, False)
        
    Exit Function

errHand:
    If ErrCenter = 1 Then
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

    objRect = mclsCommonLib.Rect
    objRect.Page = objRect.Page + 1
    mclsCommonLib.Rect = objRect
        
    NewPage = InsertPage(objRect.Page, bytCalc, strShow, blnShowPageHead, blnShowPageFoot)
        
    Exit Function

errHand:
    If ErrCenter = 1 Then
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
    
    strFile = "zlPeisGoupRpt" & Format(grsTempFile.RecordCount, "000000") & ".tmp"
    
    strFile = CreateTmpFile(strFile)
    
    GetNewTempFile = strFile
    
    grsTempFile.AddNew
    grsTempFile("文件").Value = strFile
    
End Function

Private Sub Class_Initialize()
    Set mclsCommonLib = New clsCommonLib
End Sub

Private Sub Class_Terminate()
    Set mclsCommonLib = Nothing
End Sub


