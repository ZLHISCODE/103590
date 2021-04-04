VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'######################################################################################
'##ģ �� ����cTabPara.cls
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��5��1��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    �������������ࡣ��Ӧ��TOM�ӿ��е�Para��
'##��    ����
'######################################################################################


Public Enum HAlignmentEnum
    cprHALeft = tomAlignJustify     'tomAlignLeft
    cprHACenter = tomAlignCenter
    cprHARight = tomAlignRight
End Enum

Public Enum ParaStyleEnum
    cprPSNormal = -1
    cprPSHeading1 = -2
    cprPSHeading2 = -3
    cprPSHeading3 = -4
    cprPSHeading4 = -5
    cprPSHeading5 = -6
    cprPSHeading6 = -7
    cprPSHeading7 = -8
    cprPSHeading8 = -9
    cprPSHeading9 = -10
End Enum

Public Enum ListTypeEnum
    cprLTNone = tomListNone                           '����Ŀ��
    cprLTBullet = tomListBullet                       '��Ŀ���š�
    cprLTNumberAsArabic = tomListNumberAsArabic       '���ð��������֡� (0, 1, 2, ...).
    cprLTNumberAsLCLetter = tomListNumberAsLCLetter   '����Сд��ĸ��(a, b, c, ...).
    cprLTNumberAsUCLetter = tomListNumberAsUCLetter   '���ô�д��ĸ�� (A, B, C, ...).
    cprLTNumberAsLCRoman = tomListNumberAsLCRoman     '����Сд�������֡�(i, ii, iii, ...).
    cprLTNumberAsUCRoman = tomListNumberAsUCRoman     '���ô�д�������֡� (I, II, III, ...).
    
    cprLTParenthese = tomListParentheses              '0x10000 �������ţ���: (1).
    cprLTPeriod = tomListPeriod                       '0x20000 ����һ����㡣
    cprLTPlain = tomListPlain                         '0x30000 �޷��š�
End Enum

Public Enum LineSpacingRuleEnum
    cprLSSignle = tomLineSpaceSingle        '�����о�?����Spacing��ֵ?
    cprLS1pt5 = tomLineSpace1pt5            '1.5���оࡣ����Spacing��ֵ��
    cprLSDouble = tomLineSpaceDouble        '�����о�?����Spacing��ֵ?
    cprLSAtLeast = tomLineSpaceAtLeast      '��С�о�Ϊ1�У�������ʾ��ȷֵ��
    cprLSExactly = tomLineSpaceExactly      '��ȷ�о�?
    cprLSMultiple = tomLineSpaceMultiple    '�౶�оࡣ���������㡣��1.2��ʾ�о�Ϊ1.2����׼�оࡣ
End Enum

Private m_Start As Long         'Range����ʼλ��
Private m_End As Long           'Range����ֹλ��
Private m_ReadOnly As Boolean   '�Ƿ�ֻ��
Private m_bSelection As Boolean '�Ƿ���Selection����

Private TOM As cTabTextDocument     'TOM����

Public Function GetParaFmt() As cTabParaFormat
    '��ȡ��������
    Dim R As New cTabParaFormat
    R.Alignment = Me.Alignment
    R.FirstLineIndent = Me.FirstLineIndent
    R.LeftIndent = Me.LeftIndent
    R.LineSpacing = Me.LineSpacing
    R.LineSpacingRule = Me.LineSpacingRule
    R.ListAlignment = Me.ListAlignment
    R.ListStart = Me.ListStart
    R.ListTab = Me.ListTab
    R.ListType = Me.ListType
    R.RightIndent = Me.RightIndent
    R.SpaceAfter = Me.SpaceAfter
    R.SpaceBefore = Me.SpaceBefore
    Set GetParaFmt = R
End Function

Public Function SetParaFmt(ByRef vData As cTabParaFormat)
    '���ö�������
    Me.Alignment = vData.Alignment
    Me.FirstLineIndent = vData.FirstLineIndent
    Me.LeftIndent = vData.LeftIndent
    Me.SetLineSpacing vData.LineSpacingRule, vData.LineSpacing
    Me.ListAlignment = vData.ListAlignment
    Me.ListStart = vData.ListStart
    Me.ListTab = vData.ListTab
    Me.ListType = vData.ListType
    Me.RightIndent = vData.RightIndent
    Me.SpaceAfter = vData.SpaceAfter
    Me.SpaceBefore = vData.SpaceBefore
End Function

Public Sub SetLineSpacing(ByVal Rule As LineSpacingRuleEnum, ByVal Value As Single)
    If m_ReadOnly Then Exit Sub
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.SetLineSpacing Rule, Value
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.SetLineSpacing Rule, Value
    End If
End Sub

Public Sub SetIndents(ByVal StartIndent As Single, ByVal LeftIndent As Single, ByVal RightIndent As Single)
    On Error Resume Next
    If m_ReadOnly Then Exit Sub
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.SetIndents StartIndent, LeftIndent, RightIndent
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.SetIndents StartIndent, LeftIndent, RightIndent
    End If
    Err.Clear
End Sub

Public Sub Init(ByVal cTOM As cTabTextDocument, lStart As Long, lEnd As Long, Optional bReadOnly As Boolean = False, Optional bSelection As Boolean = False)
    Set TOM = cTOM
    m_Start = lStart
    m_End = lEnd
    m_ReadOnly = bReadOnly
    m_bSelection = bSelection
End Sub

Public Property Get TabCount() As Long
    TabCount = IIf(m_bSelection, TOM.TextDocument.Selection.Para.TabCount, TOM.TextDocument.Range(m_Start, m_End).Para.TabCount)
End Property

Public Property Let WidowControl(ByVal vData As Boolean)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.WidowControl = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.WidowControl = vData
    End If
End Property

Public Property Get WidowControl() As Boolean
    WidowControl = IIf(m_bSelection, TOM.TextDocument.Selection.Para.WidowControl, TOM.TextDocument.Range(m_Start, m_End).Para.WidowControl)
End Property

Public Property Let Style(ByVal vData As ParaStyleEnum)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.Style = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.Style = vData
    End If
End Property

Public Property Get Style() As ParaStyleEnum
    Dim l As Long
    If m_bSelection Then
        l = TOM.TextDocument.Selection.Para.Style
    Else
        l = TOM.TextDocument.Range(m_Start, m_End).Para.Style
    End If
    Select Case l
    Case -1, -2, -3, -4, -5, -6, -7, -8, -9, -10
    Case Else
        l = -1
    End Select
    Style = l
End Property

Public Property Let ListAlignment(ByVal vData As HAlignmentEnum)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.ListAlignment = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.ListAlignment = vData
    End If
End Property

Public Property Get ListAlignment() As HAlignmentEnum
    ListAlignment = IIf(m_bSelection, TOM.TextDocument.Selection.Para.ListAlignment, TOM.TextDocument.Range(m_Start, m_End).Para.ListAlignment)
End Property

Public Property Let ListStart(ByVal vData As Long)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.ListStart = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.ListStart = vData
    End If
End Property

Public Property Get ListStart() As Long
    ListStart = IIf(m_bSelection, TOM.TextDocument.Selection.Para.ListStart, TOM.TextDocument.Range(m_Start, m_End).Para.ListStart)
End Property

Public Property Let ListType(ByVal vData As ListTypeEnum)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.ListType = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.ListType = vData
    End If
End Property

Public Property Get ListType() As ListTypeEnum
    Dim l As Long
    l = IIf(m_bSelection, TOM.TextDocument.Selection.Para.ListType, TOM.TextDocument.Range(m_Start, m_End).Para.ListType)
    ListType = l
End Property

Public Property Let RightIndent(ByVal vData As Single)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.RightIndent = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.RightIndent = vData
    End If
End Property

Public Property Get RightIndent() As Single
    RightIndent = IIf(m_bSelection, TOM.TextDocument.Selection.Para.RightIndent, TOM.TextDocument.Range(m_Start, m_End).Para.RightIndent)
End Property

Public Property Let LeftIndent(ByVal vData As Single)
    If m_ReadOnly Then Exit Property
    Dim lF As Long, lR As Long
    If m_bSelection Then
        lF = TOM.TextDocument.Selection.Para.FirstLineIndent
        lR = TOM.TextDocument.Selection.Para.RightIndent
    Else
        lF = TOM.TextDocument.Range(m_Start, m_End).Para.FirstLineIndent
        lR = TOM.TextDocument.Range(m_Start, m_End).Para.RightIndent
    End If
    
    If lF = tomUndefined Then lF = 0
    If vData = tomUndefined Then vData = 0
    If lR = tomUndefined Then lR = 0
           
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.SetIndents lF, vData, lR
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.SetIndents lF, vData, lR
    End If
End Property

Public Property Get LeftIndent() As Single
    LeftIndent = IIf(m_bSelection, TOM.TextDocument.Selection.Para.LeftIndent, TOM.TextDocument.Range(m_Start, m_End).Para.LeftIndent)
End Property

Public Property Let FirstLineIndent(ByVal vData As Single)
    If m_ReadOnly Then Exit Property
    Dim LL As Long, lR As Long
    If m_bSelection Then
        LL = TOM.TextDocument.Selection.Para.LeftIndent
        lR = TOM.TextDocument.Selection.Para.RightIndent
    Else
        LL = TOM.TextDocument.Range(m_Start, m_End).Para.LeftIndent
        lR = TOM.TextDocument.Range(m_Start, m_End).Para.RightIndent
    End If
    
    If vData = tomUndefined Then vData = 0
    If LL = tomUndefined Then LL = 0
    If lR = tomUndefined Then lR = 0
           
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.SetIndents vData, LL, lR
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.SetIndents vData, LL, lR
    End If
End Property

Public Property Get FirstLineIndent() As Single
    FirstLineIndent = IIf(m_bSelection, TOM.TextDocument.Selection.Para.FirstLineIndent, TOM.TextDocument.Range(m_Start, m_End).Para.FirstLineIndent)
End Property

Public Property Let SpaceAfter(ByVal vData As Single)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.SpaceAfter = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.SpaceAfter = vData
    End If
End Property

Public Property Get SpaceAfter() As Single
    SpaceAfter = IIf(m_bSelection, TOM.TextDocument.Selection.Para.SpaceAfter, TOM.TextDocument.Range(m_Start, m_End).Para.SpaceAfter)
End Property

Public Property Let SpaceBefore(ByVal vData As Single)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.SpaceBefore = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.SpaceBefore = vData
    End If
End Property

Public Property Get SpaceBefore() As Single
    SpaceBefore = IIf(m_bSelection, TOM.TextDocument.Selection.Para.SpaceBefore, TOM.TextDocument.Range(m_Start, m_End).Para.SpaceBefore)
End Property

Public Property Let Alignment(ByVal vData As HAlignmentEnum)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.Alignment = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.Alignment = vData
    End If
End Property

Public Property Get Alignment() As HAlignmentEnum
    Alignment = IIf(m_bSelection, TOM.TextDocument.Selection.Para.Alignment, TOM.TextDocument.Range(m_Start, m_End).Para.Alignment)
End Property

Public Property Get LineSpacing() As Single
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        LineSpacing = TOM.TextDocument.Selection.Para.LineSpacing
    Else
        LineSpacing = TOM.TextDocument.Range(m_Start, m_End).Para.LineSpacing
    End If
End Property

Public Property Get LineSpacingRule() As LineSpacingRuleEnum
    LineSpacingRule = IIf(m_bSelection, TOM.TextDocument.Selection.Para.LineSpacingRule, TOM.TextDocument.Range(m_Start, m_End).Para.LineSpacingRule)
End Property

Public Property Let ListTab(ByVal vData As Single)
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Para.ListTab = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Para.ListTab = vData
    End If
End Property

Public Property Get ListTab() As Single
    ListTab = IIf(m_bSelection, TOM.TextDocument.Selection.Para.ListTab, TOM.TextDocument.Range(m_Start, m_End).Para.ListTab)
End Property

