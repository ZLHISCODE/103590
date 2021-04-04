VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLisReportPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'######################################################################################################################
'接口种类： 公共接口
'实现功能： 纸张、打印机、页眉、页脚、文本段落等处理
'编制人员：
'编制时间：
'接口类名：
'接口方法： 1.InitElement
'           2.OpenSetupDialog
'           3.FormatHeadData
'           4.FormatFootData
'           5.FormatParagraph
'######################################################################################################################

Private mrsParameter        As ADODB.Recordset
Private mobjRect            As USERRECT
Private mobjFont            As USERFONT
Private mintCount           As Integer

'接口
'######################################################################################################################

Public Function SetElementParameter(ByVal rsParameter As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：设置组件属性
    '参数：rsProperty    属性记录集
    '返回：设置成功返回True,否则返回False
    '******************************************************************************************************************
    Set mrsParameter = rsParameter
    SetElementParameter = True
    
End Function

Public Function InitElement(ByRef objPaper As USERPAPER, ByRef strRegisterPath As String) As Boolean
    '******************************************************************************************************************
    '功能：公共接口：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim objclsLisPrintSet As New clsLisPrintSet
        
    InitElement = objclsLisPrintSet.GetPaper(objPaper, strRegisterPath)
    
    '固定纸张等相关参数
    '------------------------------------------------------------------------------------------------------------------
    With objPaper
        .PaperSize = 9
        .Orientation = 1
        .Width = gclsLisReportLib.GetTwipsX(21)                                         '纸张宽度
        .Height = gclsLisReportLib.GetTwipsY(29.7)                                      '纸张高度
        .BorderLeft = gclsLisReportLib.GetTwipsX(0.5)                                   '左边界
        .BorderRight = gclsLisReportLib.GetTwipsX(0.5)                                  '右边界
        .BorderTop = 0 'gclsLisReportLib.GetTwipsY(1)                                   '上边界
        .BorderBottom = gclsLisReportLib.GetTwipsY(2)                                   '下边界
        .PageHead = gclsLisReportLib.GetTwipsY(1)                                       '页眉高度
        .PageFoot = gclsLisReportLib.GetTwipsY(1)                                       '页脚高度
        .SpaceTop = gclsLisReportLib.GetTwipsY(0)                                       '
        .SpaceBottom = gclsLisReportLib.GetTwipsY(0)
    End With
    
End Function

Public Function OpenSetupDialog(frmMain As Object, ByVal strRegisterPath As String) As Boolean
    '******************************************************************************************************************
    '功能：公共接口：参数/选项设置
    '参数：
    '说明：调用此方法前，不必先调用InitData方法
    '返回：
    '******************************************************************************************************************
    Dim objPaper As USERPAPER
    
    strRegisterPath = strRegisterPath
    
    '以mm为单位保存到注册表中
    With objPaper
        .PaperSize = 9
        .Orientation = 1
        .Width = 210
        .Height = 297
        .BorderLeft = 5
        .BorderRight = 5
        .BorderTop = 10
        .BorderBottom = 10
        .PageHead = 23                               '页眉高度
        .PageFoot = 4
        .SpaceTop = 0
        .SpaceBottom = 0
    End With
    
    SaveSetting "ZLSOFT", strRegisterPath, "纸张", objPaper.PaperSize
    SaveSetting "ZLSOFT", strRegisterPath, "宽度", objPaper.Width
    SaveSetting "ZLSOFT", strRegisterPath, "高度", objPaper.Height
    SaveSetting "ZLSOFT", strRegisterPath, "纸向", objPaper.Orientation
    SaveSetting "ZLSOFT", strRegisterPath, "左边距", objPaper.BorderLeft
    SaveSetting "ZLSOFT", strRegisterPath, "右边距", objPaper.BorderRight
    SaveSetting "ZLSOFT", strRegisterPath, "上边距", objPaper.BorderTop
    SaveSetting "ZLSOFT", strRegisterPath, "下边距", objPaper.BorderBottom
    
    OpenSetupDialog = frmListReportSet.ShowDialog(frmMain, 9, False, strRegisterPath)
    
End Function

Public Function GetPageHeadData() As Boolean
    '******************************************************************************************************************
    '功能：公共接口：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strTmp                      As String
    Dim objSvrFont                  As USERFONT
    Dim varTmp                      As Variant
    Dim objRect                     As USERRECT
    Dim objFont                     As USERFONT
    Dim objclsGrID                  As clsGrid
    Dim objclsLisReportHead         As New clsLisReportHead

    On Error GoTo errHand

    '1.获取当前要输出的起始位置(下面的代码不能删除)
    '------------------------------------------------------------------------------------------------------------------
    objRect = CurrentRect

    '2.获取当前的字体设置(下面的代码不能删除)
    '------------------------------------------------------------------------------------------------------------------
    objFont = CurrentFont

    '------------------------------------------------------------------------------------------------------------------
    CurrentRect = objRect

    GetPageHeadData = True

    Exit Function

    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    gclsLisReportLib.Font = objSvrFont
End Function

Public Function GetPageFootData() As Boolean
    '******************************************************************************************************************
    '功能：公共接口：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objSvrFont              As USERFONT
    Dim objRect                 As USERRECT
    Dim objFont                 As USERFONT
    Dim objclsLisReportFoot     As New clsLisReportFoot
    
    On Error GoTo errHand
    
    '1.获取当前要输出的起始位置(下面的代码不能删除)
    '------------------------------------------------------------------------------------------------------------------
    objRect = CurrentRect
    
    '2.获取当前的字体设置(下面的代码不能删除)
    '------------------------------------------------------------------------------------------------------------------
    objFont = CurrentFont


    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - 200  '- gclsLisReportLib.Paper.PageFoot '- 200 '+ gclsLisReportLib.GetTwipsY(0.2)
    objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
    objRect.Y1 = objRect.Y0
    
    Call objclsLisReportFoot.InsertLine(objRect, objFont)
    
    objRect.X0 = gclsLisReportLib.Paper.BorderLeft
    objRect.Y0 = gclsLisReportLib.Paper.Height - gclsLisReportLib.Paper.BorderBottom - 150 '- gclsLisReportLib.Paper.PageFoot '- 150 '+ gclsLisReportLib.GetTwipsY(0.3)
    objRect.X1 = gclsLisReportLib.Paper.Width - gclsLisReportLib.Paper.BorderRight
    objRect.Y1 = 0

    objFont.HAlignment = 2
    Call objclsLisReportFoot.InsertPageNo("第 n 页 / 共 m 页", objRect, objFont)
    
    '------------------------------------------------------------------------------------------------------------------
    CurrentRect = objRect

    GetPageFootData = True
    
    Exit Function

    '------------------------------------------------------------------------------------------------------------------
errHand:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    gclsLisReportLib.Font = objSvrFont
End Function

Public Property Get CurrentRect() As USERRECT
    CurrentRect = mobjRect
End Property

Public Property Let CurrentRect(vData As USERRECT)
    mobjRect = vData
End Property

Public Property Get CurrentFont() As USERFONT
    CurrentFont = mobjFont
End Property

Public Property Let CurrentFont(vData As USERFONT)
    mobjFont = vData
End Property

Public Function FormatSetup(frmMain As Object, ByVal strRegisterPath As String) As Boolean
    '******************************************************************************************************************
    '功能：公共接口：参数/选项设置
    '参数：
    '说明：调用此方法前，不必先调用InitData方法
    '返回：
    '******************************************************************************************************************
    Dim objPaper As USERPAPER
    
    strRegisterPath = strRegisterPath
    
    '以mm为单位保存到注册表中
    With objPaper
        .PaperSize = 9
        .Orientation = 1
        .Width = 210
        .Height = 297
        .BorderLeft = 5
        .BorderRight = 5
        .BorderTop = 10
        .BorderBottom = 10
        .PageHead = 23                               '页眉高度
        .PageFoot = 4
        .SpaceTop = 0
        .SpaceBottom = 0
    End With
    
    SaveSetting "ZLSOFT", strRegisterPath, "纸张", objPaper.PaperSize
    SaveSetting "ZLSOFT", strRegisterPath, "宽度", objPaper.Width
    SaveSetting "ZLSOFT", strRegisterPath, "高度", objPaper.Height
    SaveSetting "ZLSOFT", strRegisterPath, "纸向", objPaper.Orientation
    SaveSetting "ZLSOFT", strRegisterPath, "左边距", objPaper.BorderLeft
    SaveSetting "ZLSOFT", strRegisterPath, "右边距", objPaper.BorderRight
    SaveSetting "ZLSOFT", strRegisterPath, "上边距", objPaper.BorderTop
    SaveSetting "ZLSOFT", strRegisterPath, "下边距", objPaper.BorderBottom
    
    FormatSetup = frmListReportSet.ShowDialog(frmMain, 9, False, strRegisterPath)
    
End Function

Public Function ReplaceAll(vTar As String, vFind As String, vRep As String) As String
    Dim intPos As Long

    ReplaceAll = vTar
    intPos = InStr(ReplaceAll, vFind)

    While intPos > 0
        ReplaceAll = Replace(ReplaceAll, vFind, vRep)
        intPos = InStr(ReplaceAll, vFind)
    Wend
End Function

Private Sub Class_Initialize()
    Set gclsLisReportLib = New clsLisReportLib
End Sub

Private Sub Class_Terminate()
    Set mrsParameter = Nothing
End Sub

