VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'######################################################################################
'##模 块 名：cTabSelection.cls
'##创 建 人：吴庆伟
'##日    期：2005年5月1日
'##修 改 人：
'##日    期：
'##描    述：当前选择区域。对应于TOM接口中的Selection。
'##版    本：
'######################################################################################

Option Explicit

Public Enum SelectionTypeEnum
    cprSTNone = tomNoSelection              '0   No selection and no insertion point.
    cprSTInsertPoint = tomSelectionIP       '1   Insertion point.
    cprSTNormal = tomSelectionNormal        '2   Single nondegenerate range.
    cprSTFrame = tomSelectionFrame          '3   Frame.
    cprSTColumn = tomSelectionColumn        '4   Table column.
    cprSTRow = tomSelectionRow              '5   Table rows.
    cprSTBlock = tomSelectionBlock          '6   Block selection.
    cprSTPicture = tomSelectionInlineShape  '7   Picture.
    cprSTShape = tomSelectionShape          '8   Shape.
End Enum

Public Enum HUnitEnum
    cprHUCharacter = tomCharacter
    cprHUWord = tomWord
End Enum

Public Enum VUnitEnum
    cprVULine = tomLine             '下移一行，默认值。
    cprVUParagraph = tomParagraph   '下移一段，至段首。
    cprVUScreen = tomScreen         '移动到屏幕中最末一个字符。
    cprVUWidow = tomWindow          '移动到窗口中最末一个字符。
End Enum

Private m_CanCopy As Boolean
Private m_CanPaste As Boolean
Private m_ReadOnly As Boolean   '是否只读

Private TOM As cTabTextDocument     'TOM对象

Public Property Let Text(ByVal vData As String)
    TOM.TextDocument.Selection.Text = vData
End Property

Public Property Get Text() As String
    Text = TOM.TextDocument.Selection.Text
End Property

Public Property Get Para() As cTabPara
    Dim cP As New cTabPara
    cP.Init TOM, TOM.TextDocument.Selection.Start, TOM.TextDocument.Selection.End, m_ReadOnly, True
    Set Para = cP
End Property

Public Property Get Font() As cTabFont
    Dim cF As New cTabFont
    cF.Init TOM, StartPos, EndPos, m_ReadOnly, True
    Set Font = cF
End Property

Public Sub Init(ByVal cTOM As cTabTextDocument, Optional bReadOnly As Boolean = False)
    Set TOM = cTOM
    m_ReadOnly = bReadOnly
End Sub

Public Property Get EndPos() As Long
    EndPos = TOM.TextDocument.Selection.End
End Property

Public Property Get StartPos() As Long
    StartPos = TOM.TextDocument.Selection.Start
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

Public Sub ChangeCase(Optional lStyle As CaseTypeEnum)
'用途：改变大小写。
    TOM.TextDocument.Selection.ChangeCase lStyle
End Sub

Public Sub Copy()
    TOM.TextDocument.Selection.Copy 0
End Sub

Public Sub Cut()
    TOM.TextDocument.Selection.Cut 0
End Sub

Public Sub Paste()
    TOM.TextDocument.Selection.Paste 0, 0
End Sub

Public Sub Delete()
'用途: 删除
    Dim lS As Long, lE As Long
    lS = TOM.TextDocument.Selection.Start
    lE = TOM.TextDocument.Selection.End
    If lS = lE And TOM.TextDocument.Range(lS, lS + 2) = vbCrLf And TOM.TextDocument.Range(lS, lS + 2).Font.Protected = False Then
        TOM.TextDocument.Selection.Delete tomCharacter, 2
    Else
        TOM.TextDocument.Selection.Delete tomCharacter, IIf(lS = lE, 1, 0)
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
    TOM.TextDocument.Selection.ScrollIntoView Start
End Sub

Public Sub Selected()
'用途: 将该区域设置为当前区域并高亮显示?
    TOM.TextDocument.Selection.Select
End Sub

Public Sub GetPoint(lType As GetPointTypeEnum, pX As Long, pY As Long)
'用途: 获取当前位置坐标
    Dim lX As Long, lY As Long
    TOM.TextDocument.Selection.GetPoint lType, lX, lY
    Dim pt1 As POINTAPI
    pt1.x = 0
    pt1.y = 0
    ClientToScreen TOM.hWnd, pt1
    
    pX = (lX - pt1.x) * Screen.TwipsPerPixelX + IIf(True, Screen.TwipsPerPixelX, 0)
    pY = (lY - pt1.y) * Screen.TwipsPerPixelY + IIf(True, Screen.TwipsPerPixelX, 0)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'相对于Range对象所新增的函数：
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function GetType() As SelectionTypeEnum
    GetType = TOM.TextDocument.Selection.Type
End Function

Public Sub MoveLeft(Optional Unit As HUnitEnum = cprHUCharacter, Optional Count As Long = 1, Optional Extend As Boolean = False)
' 用途：选取向左移动。
' Unit可取tomCharacter和tomWord，默认为tomCharacter；
' Count表示移动的数目，默认为1；
' Extend表示是否是扩展当前选区还是只改变插入点，True表示扩展，False表示改变插入点，默认为False。
    TOM.TextDocument.Selection.MoveLeft Unit, Count, Extend
End Sub

Public Sub MoveRight(Optional Unit As HUnitEnum = cprHUCharacter, Optional Count As Long = 1, Optional Extend As Boolean = False)
'用途：选取向右移动。
    TOM.TextDocument.Selection.MoveRight Unit, Count, Extend
End Sub

Public Sub MoveDown(Optional Unit As VUnitEnum = cprVULine, Optional Count As Long = 1, Optional Extend As Boolean = False)
'用途：选取向左移动。
'Unit可取tomLine、tomParagraph、tomScreen和tomWindow，默认为tomLine；
'Count表示移动的数目，默认为1；
'Extend表示是否是扩展当前选区还是只改变插入点，True表示扩展，False表示改变插入点，默认为False。
    TOM.TextDocument.Selection.MoveDown Unit, Count, Extend
End Sub

Public Sub MoveUp(Optional Unit As VUnitEnum = cprVULine, Optional Count As Long = 1, Optional Extend As Boolean = False)
'用途：选取向左移动。
    TOM.TextDocument.Selection.MoveUp Unit, Count, Extend
End Sub
