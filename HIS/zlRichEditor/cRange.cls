VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'######################################################################################
'##模 块 名：cRange.cls
'##创 建 人：吴庆伟
'##日    期：2005年5月1日
'##修 改 人：
'##日    期：
'##描    述：指定区域。对应于TOM接口中的Range()。
'##版    本：
'######################################################################################

Option Explicit

Public Enum CaseTypeEnum
    cprCTLowerCase = tomLowerCase         '小写.
    cprCTUpperCase = tomUpperCase         '大写.
    cprCTTitleCase = tomTitleCase         '词首大写.
    cprCTSentenceCase = tomSentenceCase   '句首大写.
    cprCTToggleCase = tomToggleCase       '取相反值.
End Enum

Public Enum StartPosEnum
    cprSPStart = tomStart           '起始位置
    cprSPEnd = tomEnd               '结束位置
End Enum

Public Enum GetPointTypeEnum
    cprGPStart = tomStart           '起始位置
    cprGPEnd = tomEnd               '结束位置
    cprGPTop = TA_TOP               '边界矩形的顶部边缘
    cprGPBaseLine = TA_BASELINE     '文本基线
    cprGPBottom = TA_BOTTOM         '边界矩形的底部边缘
    cprGPLeft = TA_LEFT             '边界矩形的左边缘
    cprGPCenter = TA_CENTER         '边界矩形的中心位置
    cprGPRight = TA_RIGHT           '边界矩形的右边缘
End Enum

Private m_CanCopy As Boolean
Private m_CanPaste As Boolean

Private m_Start As Long         'Range的起始位置
Private m_End As Long           'Range的终止位置
Private m_ReadOnly As Boolean   '是否只读

Private TOM As cTextDocument     'TOM对象

Public Property Let Text(ByVal vData As String)
    TOM.TextDocument.Range(m_Start, m_End).Text = vData
End Property

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = TOM.TextDocument.Range(m_Start, m_End).Text
End Property

Public Property Get Para() As cPara
    Dim cP As New cPara
    cP.Init TOM, m_Start, m_End, m_ReadOnly
    Set Para = cP
End Property

Public Property Get Font() As cFont
    Dim cF As New cFont
    cF.Init TOM, m_Start, m_End, m_ReadOnly
    Set Font = cF
End Property

Public Sub Init(ByVal cTOM As cTextDocument, lStart As Long, lEnd As Long, Optional bReadOnly As Boolean = False)
    Set TOM = cTOM
    m_Start = lStart
    m_End = lEnd
    m_ReadOnly = bReadOnly
End Sub

Public Property Let EndPos(vData As Long)
    m_End = vData
End Property

Public Property Get EndPos() As Long
Attribute EndPos.VB_Description = "获取当前范围结束位置。只读。"
    EndPos = m_End
End Property

Public Property Let StartPos(vData As Long)
    m_Start = vData
End Property

Public Property Get StartPos() As Long
Attribute StartPos.VB_Description = "获取当前范围起始位置。只读。"
    StartPos = m_Start
End Property

Public Property Get CanPaste() As Boolean
    CanPaste = SendMessageLong(TOM.hWnd, EM_CANPASTE, 0, 0)
End Property

Public Property Get CanCopy() As Boolean
    If TOM.TextDocument.Selection.End > TOM.TextDocument.Selection.Start Then
        CanCopy = True
    Else
        CanCopy = False
    End If
End Property

Public Sub Copy()
    TOM.TextDocument.Range(m_Start, m_End).Copy 0
End Sub

Public Sub Cut()
    TOM.TextDocument.Range(m_Start, m_End).Cut 0
End Sub

Public Sub Paste()
    TOM.TextDocument.Range(m_Start, m_End).Paste 0, 0
End Sub

Public Sub ChangeCase(Optional lStyle As CaseTypeEnum)
'用途：改变大小写。
    TOM.TextDocument.Range(m_Start, m_End).ChangeCase lStyle
End Sub

Public Sub Delete()
'用途: 删除
    If m_Start = m_End And TOM.TextDocument.Range(m_Start, m_Start + 2) = vbCrLf And TOM.TextDocument.Range(m_Start, m_Start + 2).Font.Protected = False Then
        TOM.TextDocument.Range(m_Start, m_End).Delete tomCharacter, 2
    Else
        TOM.TextDocument.Range(m_Start, m_End).Delete tomCharacter, IIf(m_Start = m_End, 1, 0)
    End If
End Sub

Public Sub Expand()
'用途: 展开提纲  用于大纲模式
    SendMessageLong TOM.hWnd, EM_OUTLINE, EMO_EXPAND, 1
End Sub

Public Sub Collapse()
'用途: 折叠提纲?用于大纲模式?
    SendMessageLong TOM.hWnd, EM_OUTLINE, EMO_EXPAND, -1
End Sub

Public Sub ScrollIntoView(Start As StartPosEnum)
'用途: 将选中区域滚动到可显示的位置?bStart可以取值为: tomEnd和tomStart?
    TOM.TextDocument.Range(m_Start, m_End).ScrollIntoView Start
End Sub

Public Sub Selected()
'用途: 将该区域设置为当前区域并高亮显示?
    TOM.TextDocument.Range(m_Start, m_End).Select
End Sub

Public Sub GetPoint(lType As GetPointTypeEnum, pX As Long, pY As Long)
'用途: 获取当前位置坐标
    Dim lX As Long, lY As Long
    TOM.TextDocument.Range(m_Start, m_End).GetPoint lType, lX, lY
    Dim pt1 As POINTAPI
    pt1.x = 0
    pt1.Y = 0
    ClientToScreen TOM.hWnd, pt1
    pX = (lX - pt1.x) * Screen.TwipsPerPixelX + IIf(True, Screen.TwipsPerPixelX, 0)
    pY = (lY - pt1.Y) * Screen.TwipsPerPixelY + IIf(True, Screen.TwipsPerPixelX, 0)
End Sub


Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Set TOM = Nothing
End Sub
