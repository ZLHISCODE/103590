VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.UserControl IDKindNew 
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   ScaleHeight     =   2460
   ScaleWidth      =   4200
   ToolboxBitmap   =   "IDKindNew.ctx":0000
   Begin VB.PictureBox picDown 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   525
      ScaleHeight     =   420
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   285
      Width           =   255
      Begin VB.Image imgDown 
         Height          =   120
         Left            =   645
         Picture         =   "IDKindNew.ctx":0312
         Top             =   810
         Width           =   120
      End
   End
   Begin VB.PictureBox picKind 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   195
      ScaleHeight     =   420
      ScaleWidth      =   330
      TabIndex        =   0
      Top             =   300
      Width           =   330
   End
   Begin VB.Shape shpWhite 
      BorderColor     =   &H80000005&
      FillColor       =   &H00404040&
      Height          =   795
      Left            =   225
      Top             =   540
      Width           =   3480
   End
   Begin VB.Shape shapBack 
      BorderColor     =   &H8000000A&
      Height          =   825
      Left            =   0
      Top             =   285
      Width           =   3510
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   105
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "IDKindNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum IDKind_Appearance
    ShowFlat = 0
    Show3D = 1
    ShowSunken3D = 2
End Enum

Public Enum IDKind_BorderStyle
    ShowNone = 0
    ShowFixed_Single = 1
End Enum

Public Enum IDKind_CaptionAlignment
    Show_Left_Justify = 0
    Show_Right_Justify = 1
    Show_Center = 2
End Enum

Public Enum Mach_Mode   '匹配方式
    CardTypeName = 0
    CardTypeID = 1
    CardTypeIndex = 2
End Enum
Public Enum IDKind_RegType
    Save_注册信息 = 0
    Save_公共全局 = 1
    Save_公共模块 = 2
    Save_私有全局 = 3
    Save_私有模块 = 4
End Enum
Private mRegType As gRegType

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private mstrPopuCaption As String
Private mstrCardType As String  '当前卡类别
Private mblnNotItemClick As Boolean
Private mcnOracle As ADODB.Connection
Private vRect As RECT
Private mobjCurCard As Card
Private mobjCards As Cards  '有效卡类别
Private mobjDefaultCard As Card '当前缺省卡类别
Private mlngDefaultCardID As Long '缺省的读卡类别ID
Private mOnlyThreeCard As Boolean

Const mMenu_Kinds = 1000
Private WithEvents mobjParent As Form
Attribute mobjParent.VB_VarHelpID = -1
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Dim WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Dim WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Dim WithEvents mobjTxtInput As TextBox
Attribute mobjTxtInput.VB_VarHelpID = -1

Private mobjPopup As CommandBarPopup
Private mobjCommandBar As CommandBar
Private mobjControl As CommandBarControl
Private mblnTextFocus As Boolean
Private mblnNotOk As Boolean
Private mblnSingle As Boolean

'缺省属性值:
Const m_def_Locked = False
Const m_def_ProductName = "IDKindNew"
Const m_def_SaveRegType = IDKind_RegType.Save_私有全局
Const m_def_OnlyReadCardNo = True

Const m_def_MustSelectItems = ""
Const m_def_AllowAutoICCard = False
Const m_def_AllowAutoIDCard = False
Const m_def_AllowAutoCommCard = True
Const m_def_NotAutoAppendKind = False
Const m_def_NotContainFastKey = "F1;CTRL+F1;F12;CTRL+F12"
Const m_def_AutoSize = False
Const m_def_KeyShift = 2
Const m_def_SmallStyle = False
Const m_def_DefaultCardType = "就诊卡"
Const m_def_ShowPropertySet = False
Const m_def_IDKind = 0
Const m_def_CaptionAlignment = 2
'Const m_def_AutoSize = False
Const m_def_BorderStyle = IDKind_BorderStyle.ShowNone
Const m_def_IDKindStr = "姓|姓名或就诊卡|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;手|手机号|0"
Const m_def_Appearance = 0
Const m_def_ShowSortName = True
'属性变量:
Dim m_Locked As Boolean
Dim m_ProductName As String
Dim m_SaveRegType As IDKind_RegType
Dim m_OnlyReadCardNo As Boolean
Dim m_MustSelectItems As String
Dim m_AllowAutoICCard As Boolean
Dim m_AllowAutoIDCard As Boolean
Dim m_AllowAutoCommCard As Boolean
Dim m_NotContainFastKey As String
Dim m_AutoSize As Boolean
Dim m_NotAutoAppendKind As Boolean
Dim m_KeyShift As Long
Dim m_SmallStyle As Boolean
Dim m_DefaultCardType As String
Dim m_ShowPropertySet As Boolean
Dim m_IDKind As Integer
Dim m_CaptionAlignment As Integer
'Dim m_AutoSize As Boolean
Dim m_BorderStyle As Integer
Dim m_Cards As Cards
Dim m_IDKindStr As String
Dim m_Appearance As Byte
Dim m_ShowSortName As Boolean
'事件声明:
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event Click(objCard As Object)
Public Event ItemClick(Index As Integer, objCard As Object)
Public Event ReadCard(ByVal objCard As Object, objPatiInfor As Object, blnCancel As Boolean)
Private mobjPubOneCard As clsPublicOneCard

Private Sub InitCardsObject()
    Dim strValue As String
    Dim objCard As Card, strCardTypes As String
    Dim i As Long, bln加密显示 As Boolean
    Dim DefaultCardType As Long
    
    On Error GoTo errHandle
    
    Set mobjCards = New Cards
    If Cards Is Nothing Then Exit Sub
    
     
    If mobjPubOneCard Is Nothing Then Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
 
    '78768:李南春,2014/11/26,将快键保存到参数表
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("向前滚动-功能键", glngSys, 1153)
    End If
    
    If strValue = "" Then Call GetRegInFor(mRegType, "医疗卡类别", "向前滚动-功能键", strValue)
    Cards.向前滚动功能键 = strValue
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("向前滚动-快键", glngSys, 1153)
    End If
    If strValue = "" Then Call GetRegInFor(mRegType, "医疗卡类别", "向前滚动-快键", strValue)
    If strValue = "" Then Cards.向前滚动功能键 = "SHIFT"
    Cards.向前滚动快键 = IIf(strValue = "", "F4", strValue)
    mobjCards.向前滚动功能键 = Cards.向前滚动功能键
    mobjCards.向前滚动快键 = Cards.向前滚动快键
    
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("向后滚动-功能键", glngSys, 1153)
    End If
    If strValue = "" Then Call GetRegInFor(mRegType, "医疗卡类别", "向后滚动-功能键", strValue)
    Cards.向后滚动功能键 = strValue
    mobjCards.向后滚动功能键 = Cards.向后滚动功能键
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("向后滚动-快键", glngSys, 1153)
    End If
    If strValue = "" Then Call GetRegInFor(mRegType, "医疗卡类别", "向后滚动-快键", strValue)
    Cards.向后滚动快键 = IIf(strValue = "", "F4", strValue)
    mobjCards.向后滚动快键 = Cards.向后滚动快键
    
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("读卡-功能键", glngSys, 1153)
    End If
    If strValue = "" Then Call GetRegInFor(mRegType, "医疗卡类别", "读卡-功能键", strValue)
    Cards.读卡功能键 = strValue
    mobjCards.读卡功能键 = Cards.读卡功能键
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("读卡-快键", glngSys, 1153)
    End If
    If strValue = "" Then Call GetRegInFor(mRegType, "医疗卡类别", "读卡-快键", strValue)
    Cards.读卡快键 = IIf(strValue = "", "空格键", strValue)
    mobjCards.读卡快键 = Cards.读卡快键
    
    '78768:李南春,2014/11/26,缺省读卡类别
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("缺省读卡类别", glngSys, 1153, 0)
    Else
        Call GetRegInFor(mRegType, "医疗卡类别", "缺省读卡类别", strValue)
    End If
    mlngDefaultCardID = Val(strValue)
    
    Set mobjDefaultCard = Nothing
    With mobjCards
        For i = 1 To Cards.Count
            If m_ShowPropertySet And Cards(i).启用 Then
                Call GetRegInFor(mRegType, "医疗卡类别\" & Cards(i).名称, "启用", strValue)
                If strValue = "" Then   '缺省启用
                    Cards(i).启用 = True
                Else
                    Cards(i).启用 = Val(strValue) <> 0
                End If
            ElseIf Cards(i).接口序号 = 0 Then
                Call GetRegInFor(mRegType, "医疗卡类别\" & Cards(i).名称, "启用", strValue)
                If strValue = "" Then   '缺省启用
                    Cards(i).启用 = True
                Else
                    Cards(i).启用 = Val(strValue) <> 0
                End If
            End If
            
            Call GetRegInFor(mRegType, "医疗卡类别\" & Cards(i).名称, "读卡-功能键", strValue)
            Cards(i).功能键 = strValue
            Call GetRegInFor(mRegType, "医疗卡类别\" & Cards(i).名称, "读卡-快键", strValue)
            Cards(i).快键 = strValue
            '缺省的医疗卡类别
            '118959:李南春，2018/1/3,缺省规则调整
            '1.以医疗卡类别.缺省标志=1的医疗卡为缺省类别
            '2.以IDKindNew的DefaultCardType为缺省类别
            '3.以第一个启用的医疗卡类别为缺省医疗卡类别
            '76843:李南春,2014/8/22,设置缺省的医疗卡对象
            If Cards(i).缺省标志 Then
                DefaultCardType = Cards(i).接口序号
                Set mobjDefaultCard = Cards(i)
            End If
            If m_DefaultCardType <> "" And mobjDefaultCard Is Nothing And (Cards(i).名称 = m_DefaultCardType _
                Or Cards(i).接口序号 = Val(m_DefaultCardType) And Val(m_DefaultCardType) <> 0) Then
                DefaultCardType = Cards(i).接口序号
                Set mobjDefaultCard = Cards(i)
            End If
            If Cards(i).接口序号 > 0 And objCard Is Nothing Then
                If objCard Is Nothing Then
                    DefaultCardType = Cards(i).接口序号
                    Set objCard = Cards(i)
                End If
            End If
            
            If Cards(i).名称 Like "姓名*" Then
                Cards(i).模糊查找项 = True
            End If
            If Cards(i).启用 Then
                If Not bln加密显示 Then bln加密显示 = IIf(Cards(i).卡号密文规则 <> "" And Cards(i).卡号密文规则 <> "0", True, False)
                If Cards(i).是否模糊查找 And Cards(i).接口序号 > 0 Then
                    strCardTypes = strCardTypes & "," & Cards(i).接口序号
                End If
                If Cards(i).接口序号 = 0 Then
                    mobjCards.Add Cards(i), "M" & Cards(i).名称
                Else
                    mobjCards.Add Cards(i), "K" & Cards(i).接口序号
                End If
            End If
        Next
        .按缺省卡查找 = Cards.按缺省卡查找
        If strCardTypes <> "" Then strCardTypes = Mid(strCardTypes, 2)
        .模糊查找类别 = strCardTypes
        .加密显示 = bln加密显示
        If Not gobjCards Is Nothing Then
            gobjCards.加密显示 = bln加密显示
        End If
        If mobjDefaultCard Is Nothing Then Set mobjDefaultCard = objCard
        m_DefaultCardType = DefaultCardType
    End With
    Call CreatePopuMenu
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ReInitCards()
    Call InitPopuCationVar
    Call SetCaption
End Sub
Private Sub InitPopuCationVar()
   Dim i As Integer, strCardType As String
    
    mstrPopuCaption = IIf(ShowSortName, "姓", "姓名")
    m_IDKind = -1
    If mobjCards Is Nothing Then Exit Sub
    If mobjCards.Count = 0 Then Exit Sub
    
    Err = 0: On Error Resume Next
    If mstrCardType = "" Then
        If mobjCards.Count <> 0 Then
            mstrCardType = IIf(mobjCards(1).接口序号 > 0, mobjCards(1).接口序号, mobjCards(1).名称)
        End If
    End If
    For i = 1 To mobjCards.Count
        strCardType = IIf(mobjCards(i).接口序号 > 0, mobjCards(i).接口序号, mobjCards(i).名称)
        If strCardType = mstrCardType Then
            Set mobjCurCard = mobjCards(i): m_IDKind = i
            Exit For
        End If
    Next
    
    If mobjCurCard Is Nothing Then
        If mobjCards.Count <> 0 Then
            Set mobjCurCard = mobjCards(1)
        End If
        If mobjCurCard Is Nothing Then mstrPopuCaption = "": Exit Sub
    End If
    
    mstrPopuCaption = IIf(ShowSortName, mobjCurCard.短名, mobjCurCard.名称)
    If mstrCardType <> IIf(mobjCurCard.接口序号 > 0, mobjCurCard.接口序号, mobjCurCard.名称) Then
        strCardType = IIf(mobjCurCard.接口序号 > 0, mobjCurCard.接口序号, mobjCurCard.名称)
    End If
    picKind.ToolTipText = mobjCurCard.名称
    If Ambient.UserMode And mblnNotItemClick = False Then
        RaiseEvent ItemClick(m_IDKind, mobjCurCard)
    End If
 End Sub
Private Sub SetCaption()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置标题
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call SetKindWidth
    Call RePicKindStatu
End Sub
Private Sub RePicKindStatu()
    Dim intType As Integer
    If Appearance = ShowSunken3D Then
        UserControl.BorderStyle = 1
        If mobjCurCard Is Nothing Then
            intType = 0
        '78768:李南春,2014/11/26,缺省读卡类别
        '85565,李南春,2015/7/19:读卡性质,接触式读卡需要点击
        ElseIf CheckReadCard Then
            intType = 2
        ElseIf mobjCurCard.是否接触式读卡 Then
            intType = 2
        Else
            intType = 0
        End If
        SetCommandStatu intType, 0
        SetCommandStatu 2, 1
        Exit Sub
    End If
    UserControl.BorderStyle = 0
    Select Case Appearance
    Case ShowFlat
         intType = 0
    Case Else
         intType = 1
    End Select
    SetCommandStatu intType
End Sub
Private Sub SetKindWidth()
    Dim lngSkip As Long '
    '设置Kind的宽度
    If Not AutoSize Then
        '不自动调整
        lngSkip = 0
        If Not (BorderStyle = ShowNone Or Appearance <> 0) Then
            lngSkip = 90
        End If
        If ScaleWidth - picDown.Width - lngSkip < 0 Then
            picKind.Width = 0
        Else
            picKind.Width = ScaleWidth - picDown.Width - lngSkip
        End If
        Exit Sub
    End If
    
    picKind.Width = picKind.TextWidth(mstrPopuCaption) + 120
    If BorderStyle = ShowNone Or Appearance <> ShowFlat Then
        UserControl.Width = picKind.Width + picDown.Width + IIf(Appearance = ShowSunken3D, 80, 0)
        Exit Sub
    End If
    UserControl.Width = picKind.Width + picDown.Width + 80
End Sub


Private Sub zlCommandBarDef()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定义类
    '编制:刘兴洪
    '日期:2012-08-15 15:51:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsThis.DeleteAll
    Exit Sub
errHandle:
    MsgBox Err.Description
End Sub
Private Sub SetCommandStatu(ByVal intStyle As Integer, Optional intType As Integer)
    Dim intAlignMent As gAlignment
    If CaptionAlignment = Show_Center Then
        intAlignMent = gAlignment.mCenterAgnmt
    ElseIf CaptionAlignment = Show_Left_Justify Then
        intAlignMent = gAlignment.mLeftAgnmt
    Else
        intAlignMent = gAlignment.mRightAgnmt
    End If
    'intType:0-所有;1-OnlyDown;2-OnlyKind
    If intType = 0 Or intType = 2 Then
        zlRaisEffect picKind, intStyle, mstrPopuCaption, intAlignMent
    End If
    If intType = 0 Or intType = 1 Then
        zlRaisEffect picDown, intStyle, " ", intAlignMent
    End If
End Sub
Private Sub CreatePopuMenu()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建临时菜单
    '编制:刘兴洪
    '日期:2012-11-21 09:49:35
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim j As Long, strCardType As String
    Dim objCurCard As Card
    
    If mobjCards Is Nothing Then
       Call InitCardsObject: ReInitCards
    End If
    If mobjCards Is Nothing Then Exit Sub
    Set mobjCommandBar = cbsThis.Add("PopupPati", xtpBarPopup)
      
    j = 1
    With mobjCommandBar.Controls
        For Each objCard In mobjCards
            Set mobjControl = .Add(xtpControlButton, mMenu_Kinds + j, objCard.名称)
            strCardType = IIf(objCard.接口序号 > 0, objCard.接口序号, objCard.名称)
            
            mobjControl.Parameter = strCardType
            If mstrCardType = strCardType Then
                mstrPopuCaption = IIf(ShowSortName, objCard.短名, objCard.名称)
                mstrCardType = strCardType
            End If
            j = j + 1
        Next
        
        If ShowPropertySet Then
            '显示属性设置
            Set mobjControl = .Add(xtpControlButton, mMenu_Kinds + j, "类别属性设置")
            mobjControl.Parameter = "PropertySet"
            mobjControl.BeginGroup = True
        End If

    End With
End Sub
Private Sub AddPopu(ByVal X As Long, ByVal Y As Long)
    vRect = GetControlRect(picKind.hWnd)
    vRect.Left = vRect.Left - 2
    vRect.Top = vRect.Top + 2
    Call CreatePopuMenu
    If Not mobjCommandBar Is Nothing Then Call mobjCommandBar.ShowPopup(, vRect.Left, vRect.Top + picDown.Height)
End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Parameter = "PropertySet" Then
        '属性设置
        If frmIDKindSet.ShowSetWin(mobjParent, mcnOracle, Cards, mRegType, NotContainFastKey, MustSelectItems) = False Then
             shapBack.BorderColor = &H8000000A
             picDown.Tag = ""
             picKind.Tag = ""
            Call SetCaption
            Exit Sub
        End If
        Call InitCardsObject
        Call ReInitCards
         shapBack.BorderColor = &H8000000A
         picDown.Tag = ""
         picKind.Tag = ""
        Call SetCaption
        Exit Sub
    End If
    mstrCardType = Control.Parameter
    Call ReInitCards
    RaiseEvent ItemClick(m_IDKind, mobjCurCard)
    shapBack.BorderColor = &H8000000A
     picDown.Tag = ""
     picKind.Tag = ""
    Call SetCaption
'    DoEvents
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.Parameter = mstrCardType Then
        Control.Checked = True
    End If
    Control.Enabled = Not Locked
End Sub

Private Sub imgDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If Appearance = ShowSunken3D Then
        Call SetCommandStatu(-1, 1)
    Else
        Call SetCommandStatu(-1)
    End If
    Call AddPopu(X, Y)
End Sub

Private Sub imgDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
  Select Case Appearance
    Case ShowFlat
        Call SetCommandStatu(0)
    Case ShowSunken3D
        Call SetCommandStatu(0, 0)
        'Call SetCommandStatu(1, 1)
    Case Else
        Call SetCommandStatu(1)
    End Select
    shapBack.BorderColor = &H8000000A
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    Dim objPati As clsPatiInfor, strErrMsg As String
    Dim objCard As Card
    Set objCard = GetIDKindCard(strCardType, CardTypeName)
    If objCard Is Nothing Then Exit Sub
    
    Set objPati = New clsPatiInfor
    If strXmlCardInfor <> "" Then
       Call zlGetPatiInforFromXML(mcnOracle, strXmlCardInfor, objPati, strErrMsg)
    End If
    
    If objPati Is Nothing Then Set objPati = New clsPatiInfor
    objPati.卡号 = strCardNo
    
    RaiseEvent ReadCard(objCard, objPati, blnCancel)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    Dim objPati As clsPatiInfor
    Dim blnCancel As Boolean
    Dim objCard As Card
    Set objCard = GetIDKindCard("IC卡", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    Set objPati = New clsPatiInfor
    objPati.卡号 = strCardNo
    RaiseEvent ReadCard(objCard, objPati, blnCancel)
End Sub
Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    Dim objPati As clsPatiInfor
    Dim blnCancel As Boolean
    Dim objCard As Card
    Dim objStdPic As StdPicture
    Set objCard = GetIDKindCard("身份证", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    Set objPati = New clsPatiInfor
    objPati.卡号 = strID
    objPati.姓名 = strName
    objPati.性别 = strSex
    objPati.出生日期 = Format(datBirthday, "yyyy-mm-DD HH:MM:SS")
    objPati.出生地址 = strAddress
    objPati.身份证号 = strID
    objPati.民族 = strNation
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    objPati.照片 = objStdPic
    objPati.照片文件 = ""
    RaiseEvent ReadCard(objCard, objPati, blnCancel)
End Sub

Private Function IDKindClick(ByVal objCard As Card, Optional ByVal blnFastKeyUse As Boolean = False) As Boolean
    Dim objPati As New clsPatiInfor
    Dim blnCancel As Boolean, strErrMsg As String
    
    Dim strExpand As String, strOutCardNO As String, strOutPatiInforXML As String, strPhotoFile As String
    '点击IDkindClick
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        '72861:刘尔旋,2014-05-09,快捷键无法调用IC卡的读卡问题
        If mobjICCard Is Nothing Then
            'If blnFastKeyUse = True Then   '
                RaiseEvent Click(mobjCurCard)
                Call SetCaption
                Call ClearTag
                IDKindClick = True
            'End If
            Exit Function
        End If
        objPati.卡号 = mobjICCard.Read_Card()
        RaiseEvent ReadCard(objCard, objPati, blnCancel)
        If blnCancel = True Then Exit Function
        IDKindClick = True
        Exit Function
    End If
    Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    
    If objCard.接口序号 <= 0 Or mobjPubOneCard Is Nothing Then Exit Function
    
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOnlyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOnlyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败
    If mobjPubOneCard.objThirdSwap.zlReadCard(mobjParent, glngModul, objCard.接口序号, m_OnlyReadCardNo, strExpand, strOutCardNO, strOutPatiInforXML, strPhotoFile) = False Then Exit Function
    If Not m_OnlyReadCardNo Then
        Call zlGetPatiInforFromXML(mcnOracle, strOutPatiInforXML, objPati, strErrMsg)
        If objPati.照片 Is Nothing And Trim(strPhotoFile) <> "" Then
            On Error Resume Next
            objPati.照片 = LoadPicture(strPhotoFile)
            Err = 0: On Error GoTo 0
        End If
    End If
    If objPati Is Nothing Then Set objPati = New clsPatiInfor
    If objPati.卡号 = "" Then objPati.卡号 = strOutCardNO
    RaiseEvent ReadCard(objCard, objPati, blnCancel)
    If blnCancel = True Then Exit Function
    IDKindClick = True
End Function

 

Private Sub mobjTxtInput_Change()
    If mobjTxtInput.Locked Or mobjTxtInput.Visible = False Or mobjTxtInput.Enabled = False Then
        Exit Sub
    End If
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(mobjTxtInput.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(mobjTxtInput.Text = "")
    Call SetAutoReadCard(mobjTxtInput.Text = "")
End Sub

Private Sub mobjTxtInput_GotFocus()
    If mobjTxtInput.Locked Then Exit Sub
    mblnTextFocus = True
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(mobjTxtInput.Text = "")
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(mobjTxtInput.Text = "")
    Call SetBrushCardObject(True)
    Call SetAutoReadCard(mobjTxtInput.Text = "")
End Sub
Private Sub mobjTxtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If mobjTxtInput.Locked Or mobjTxtInput.Enabled = False Then Exit Sub
    If mobjTxtInput.Text <> "" Then Exit Sub
    If ActiveFastKeyInside = True Then Exit Sub
End Sub

Private Sub mobjTxtInput_KeyPress(KeyAscii As Integer)
    If m_Cards Is Nothing Then Exit Sub
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    '问题:51488
    If Not mobjCurCard Is Nothing Then
         If mobjCurCard.是否刷卡 Or mobjCurCard.是否扫描 Then Exit Sub
    End If
    If (m_Cards.读卡快键 = "空格键" Or m_Cards.读卡快键 = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub
End Sub
 

Private Sub mobjTxtInput_LostFocus()
    mblnTextFocus = False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    Call SetBrushCardObject(False)
    Call SetAutoReadCard(False)
End Sub

Private Sub picDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If Appearance = ShowSunken3D Then
        Call SetCommandStatu(-1, 1)
    Else
        Call SetCommandStatu(-1)
    End If
    Call AddPopu(X, Y)
    ClearTag
End Sub
Private Sub picDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Appearance = ShowSunken3D Then Exit Sub
    
   If picDown.Tag = "In" Then
        If X < 0 Or Y < 0 Or X > picDown.Width Or Y > picDown.Height Then
            picDown.Tag = ""
            ReleaseCapture
            shapBack.BorderColor = &H8000000A
            Select Case Appearance
              Case ShowFlat
                   SetCommandStatu 0
              Case Else
              End Select
        End If
    Else
        picDown.Tag = "In"
        SetCapture picDown.hWnd
        shapBack.BorderColor = vbBlue
        Select Case Appearance
          Case ShowFlat
               SetCommandStatu 1
          Case ShowSunken3D
          Case Else
          End Select
         
    End If
End Sub

Private Sub picDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Appearance
    Case ShowFlat
        Call SetCommandStatu(0)
    Case ShowSunken3D
        Call SetCommandStatu(0, 0)
        'Call SetCommandStatu(2, 1)
    Case Else
        Call SetCommandStatu(1)
    End Select
    UserControl.Tag = "": shapBack.BorderColor = &H8000000A
End Sub

Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With imgDown
        .Left = (picDown.ScaleWidth - .Width) \ 2
        .Top = (picDown.ScaleHeight - .Height) \ 2
    End With
End Sub

Private Sub picKind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If mobjCurCard Is Nothing Then Exit Sub
    '85565:李南春,2015/7/23,读卡性质
    '78768:李南春,2014/11/26,缺省读卡类别
    If mobjCurCard.是否接触式读卡 = False And CheckReadCard = False Then Exit Sub
    SetCommandStatu -1, 2
End Sub

Private Sub picKind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Appearance = ShowSunken3D Then Exit Sub
    
    If picKind.Tag = "In" Then
        If X < 0 Or Y < 0 Or X > picKind.Width Or Y > picKind.Height Then
            picKind.Tag = ""
            ReleaseCapture
            shapBack.BorderColor = &H8000000A
            If Appearance = ShowFlat Then
                SetCommandStatu 0
            ElseIf Appearance = ShowSunken3D Then
                SetCommandStatu 0, 0
                SetCommandStatu 2, 1
            End If
        End If
    Else
        picKind.Tag = "In"
        SetCapture picKind.hWnd
        shapBack.BorderColor = vbBlue
        Select Case Appearance
        Case ShowFlat
             SetCommandStatu 1
        Case ShowSunken3D
            If mobjCurCard Is Nothing Then
                SetCommandStatu 0, 0
            '78768:李南春,2014/11/26,缺省读卡类别
            ElseIf CheckReadCard Then
                SetCommandStatu 2, 0
            ElseIf mobjCurCard.是否接触式读卡 Then
                SetCommandStatu 2, 0
            Else
                SetCommandStatu 0, 0
            End If
            SetCommandStatu 2, 1
        Case Else
        End Select
    End If
End Sub

Private Sub picKind_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objCurCard As Card, bln缺省读卡 As Boolean
    Dim lngPreIDKind As Long, intIndex As Integer
    If Button <> 1 Then Exit Sub
    If mobjCurCard Is Nothing Then Exit Sub
    '78768:李南春,2014/11/26,缺省读卡类别
    bln缺省读卡 = CheckReadCard
    If mobjCurCard.是否接触式读卡 = False And bln缺省读卡 = False Then Exit Sub
    If bln缺省读卡 Then
        Set objCurCard = mobjCurCard
        Set mobjCurCard = Cards("K" & mlngDefaultCardID)
    End If
    If IDKindClick(mobjCurCard) = True Then
        If bln缺省读卡 Then Set mobjCurCard = objCurCard
        Call SetCaption
        Call ClearTag
        Exit Sub
    End If
'    RaiseEvent Click(mobjCurCard)
    If bln缺省读卡 Then Set mobjCurCard = objCurCard
    Call SetCaption
    Call ClearTag
End Sub
Private Sub ClearTag()
     shapBack.BorderColor = &H8000000A
     picKind.Tag = "": picDown.Tag = ""
End Sub

Private Sub UserControl_ExitFocus()
    '平面,恢复平面
    Select Case Appearance
    Case ShowFlat
         SetCommandStatu 0
    Case ShowSunken3D
        Call RePicKindStatu
    Case Else
        Call RePicKindStatu
     End Select
End Sub

Private Sub UserControl_Initialize()
    Call zlCommandBarDef
    glngInstanceCount = glngInstanceCount + 1
End Sub
Private Sub UserControl_Terminate()

    Err = 0: On Error Resume Next
    glngInstanceCount = IIf(glngInstanceCount > 0, glngInstanceCount - 1, 0)
    Call CloseComEvntsObject
    
    Call zlReleaseResources
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If Not mobjCurCard Is Nothing Then Set mobjCurCard = Nothing
    If Not mobjCards Is Nothing Then Set mobjCards = Nothing
    If Not mobjDefaultCard Is Nothing Then Set mobjDefaultCard = Nothing
    If Not mobjParent Is Nothing Then Set mobjParent = Nothing
    If Not mobjCommEvents Is Nothing Then Set mobjCommEvents = Nothing
    If Not mobjIDCard Is Nothing Then Set mobjIDCard = Nothing
    If Not mobjICCard Is Nothing Then Set mobjICCard = Nothing
    If Not mobjTxtInput Is Nothing Then Set mobjTxtInput = Nothing
    If Not mobjPopup Is Nothing Then Set mobjPopup = Nothing
    If Not mobjCommandBar Is Nothing Then Set mobjCommandBar = Nothing
    If Not mobjControl Is Nothing Then Set mobjControl = Nothing
    If Not mobjPubOneCard Is Nothing Then Set mobjPubOneCard = Nothing
    
End Sub

Private Sub SetCtrlVisible()
    Dim blnVisble As Boolean
    blnVisble = Appearance = 0 And BorderStyle = ShowFixed_Single
    shapBack.Visible = blnVisble
    shpWhite.Visible = blnVisble
End Sub

Private Sub UserControl_Paint()
    '打字
      
End Sub

Private Sub UserControl_Resize()

    Err = 0: On Error Resume Next
   ' If mblnNotOk Then Exit Sub
    mblnNotOk = True
    Call SetKindWidth
    With shapBack
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
        shpWhite.Top = .Top + 20
        shpWhite.Left = .Left + 20
        shpWhite.Width = .Width - 40
        shpWhite.Height = .Height - 40
    End With
    
    If Appearance = 0 Then
        
        With picDown
            .Left = ScaleWidth - .Width - IIf(shpWhite.Visible, 10, 0)
            .Top = IIf(shpWhite.Visible, shpWhite.Top + 10, ScaleTop)
            .Height = IIf(shpWhite.Visible, shpWhite.Height - 30, ScaleHeight)
        End With
        
        With picKind
            .Left = IIf(shpWhite.Visible, shpWhite.Left + 10, ScaleLeft)
            If mblnSingle = True Then .Width = ScaleWidth - .Left
            .Top = picDown.Top
            .Height = picDown.Height
        End With

        Call RePicKindStatu
        Exit Sub
    End If
    
    With picDown
    
        .Left = ScaleWidth - .Width
        .Top = UserControl.ScaleTop
        .Height = ScaleHeight
    End With
    With picKind
        .Left = ScaleLeft
        If mblnSingle = True Then .Width = ScaleWidth
        .Top = picDown.Top
        .Height = picDown.Height
    End With
    Call RePicKindStatu
    mblnNotOk = False
    'SetCommandStatu (1)
     
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get ShowSortName() As Boolean
    ShowSortName = m_ShowSortName
End Property

Public Property Let ShowSortName(ByVal New_ShowSortName As Boolean)
    m_ShowSortName = New_ShowSortName
    PropertyChanged "ShowSortName"
    Call ReInitCards
    'Call UserControl_Resize
End Property

Public Function IsMobileNo(ByVal strInput As String, Optional ByRef strRutType As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '功能:判断传入的是否为手机号
    '入参:strInput-手机号
    '出参:strRutType-查询结果:0-成功;1-不是有效号段;2-号码长度不对
    '返回:True-传入号码为手机号;False-传入号码不为手机号
    '编制:刘尔旋
    '日期:2017-1-25
    '---------------------------------------------------------------------------------------------
    strRutType = 0
    'If mcnOracle Is Nothing Then IsMobileNo = False: Exit Function
    Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    IsMobileNo = mobjPubOneCard.zlIsMobileNo(strInput, strRutType)
    Exit Function
errHand:
    strRutType = 1
End Function


'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_ShowSortName = m_def_ShowSortName
    m_Appearance = m_def_Appearance
    m_IDKindStr = m_def_IDKindStr
    m_BorderStyle = m_def_BorderStyle
'    m_AutoSize = m_def_AutoSize
    m_CaptionAlignment = m_def_CaptionAlignment
    m_IDKind = m_def_IDKind
    m_ShowPropertySet = m_def_ShowPropertySet
    m_DefaultCardType = m_def_DefaultCardType
    m_SmallStyle = m_def_SmallStyle
    m_KeyShift = m_def_KeyShift
    m_AutoSize = m_def_AutoSize
    m_NotAutoAppendKind = m_def_NotAutoAppendKind
    m_NotContainFastKey = m_def_NotContainFastKey
    m_AllowAutoICCard = m_def_AllowAutoICCard
    m_AllowAutoIDCard = m_def_AllowAutoIDCard
    m_AllowAutoCommCard = m_def_AllowAutoCommCard
    m_MustSelectItems = m_def_MustSelectItems
    m_OnlyReadCardNo = m_def_OnlyReadCardNo
    m_SaveRegType = m_def_SaveRegType
    m_ProductName = m_def_ProductName
    gstrSaveRegProceName = m_ProductName
    Call FromRegType(m_SaveRegType)
    m_Locked = m_def_Locked
    
End Sub
Private Sub FromRegType(ByVal intReg As IDKind_RegType)
    Select Case intReg
    Case Save_公共模块
        mRegType = g公共模块
    Case Save_公共全局
        mRegType = g公共全局
    Case Save_私有模块
        mRegType = g私有模块
    Case Save_私有全局
        mRegType = g私有全局
    Case Save_注册信息
        mRegType = g注册信息
    Case Else
        mRegType = g私有全局
    End Select
End Sub
'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ShowSortName = PropBag.ReadProperty("ShowSortName", m_def_ShowSortName)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_IDKindStr = PropBag.ReadProperty("IDKindStr", m_def_IDKindStr)
    m_NotAutoAppendKind = PropBag.ReadProperty("NotAutoAppendKind", m_def_NotAutoAppendKind)
    Set m_Cards = zlGetKindCards(m_IDKindStr, , NotAutoAppendKind, mOnlyThreeCard)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
'    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    m_CaptionAlignment = PropBag.ReadProperty("CaptionAlignment", m_def_CaptionAlignment)
    
    Set picDown.Font = PropBag.ReadProperty("Font", Ambient.Font)
    picKind.FontBold = PropBag.ReadProperty("FontBold", 0)
    picKind.FontSize = PropBag.ReadProperty("FontSize", 9)
    picKind.FontName = PropBag.ReadProperty("FontName", UserControl.FontName)
    picKind.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set UserControl.Font = picDown.Font
    UserControl.FontBold = picKind.FontBold
    UserControl.FontSize = picKind.FontSize
    UserControl.FontName = picKind.FontName
    UserControl.ForeColor = picKind.ForeColor
    
    m_IDKind = PropBag.ReadProperty("IDKind", m_def_IDKind)
    m_ShowPropertySet = PropBag.ReadProperty("ShowPropertySet", m_def_ShowPropertySet)
    m_DefaultCardType = PropBag.ReadProperty("DefaultCardType", m_def_DefaultCardType)
    m_SmallStyle = PropBag.ReadProperty("SmallStyle", m_def_SmallStyle)
    m_KeyShift = PropBag.ReadProperty("KeyShift", m_def_KeyShift)
   
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    picDown.Enabled = UserControl.Enabled
    picDown.Enabled = UserControl.Enabled
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
     
    m_NotContainFastKey = PropBag.ReadProperty("NotContainFastKey", m_def_NotContainFastKey)
    m_AllowAutoICCard = PropBag.ReadProperty("AllowAutoICCard", m_def_AllowAutoICCard)
    m_AllowAutoIDCard = PropBag.ReadProperty("AllowAutoIDCard", m_def_AllowAutoIDCard)
    m_AllowAutoCommCard = PropBag.ReadProperty("AllowAutoCommCard", m_def_AllowAutoCommCard)
    m_MustSelectItems = PropBag.ReadProperty("MustSelectItems", m_def_MustSelectItems)
    m_OnlyReadCardNo = PropBag.ReadProperty("OnlyReadCardNo", m_def_OnlyReadCardNo)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    picDown.BackColor = UserControl.BackColor
    picKind.BackColor = UserControl.BackColor
    m_SaveRegType = PropBag.ReadProperty("SaveRegType", m_def_SaveRegType)
    m_ProductName = PropBag.ReadProperty("ProductName", m_def_ProductName)
    gstrProductName = m_ProductName
    Call FromRegType(m_SaveRegType)

    Call setControlParentFormObject
    Call SetCtrlVisible
    Call ReInitCards
    If Ambient.UserMode Then
        If m_DefaultCardType = "" Then
            Set mobjDefaultCard = Nothing
        ElseIf IsNumeric(m_DefaultCardType) Then
            Set mobjDefaultCard = GetIDKindCard(m_DefaultCardType, CardTypeID)
        Else
            Set mobjDefaultCard = GetIDKindCard(m_DefaultCardType, CardTypeName)
        End If
    End If
    
    m_Locked = PropBag.ReadProperty("Locked", m_def_Locked)
End Sub
Private Sub setControlParentFormObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置当前控件的父窗体对象
    '编制:刘兴洪
    '日期:2018-12-19 10:02:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
   If Not Ambient.UserMode Or UCase(TypeName(UserControl.Parent)) = UCase("PatiIdentify") Then Exit Sub
   Err = 0: On Error Resume Next
   If TypeOf UserControl.Parent Is Form Then Set mobjParent = UserControl.Parent
   Err.Clear: On Error GoTo 0
End Sub
 
 
Private Sub UserControl_Show()
   Err = 0: On Error Resume Next
 
    Call UserControl_Resize
End Sub



'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ShowSortName", m_ShowSortName, m_def_ShowSortName)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("IDKindStr", m_IDKindStr, m_def_IDKindStr)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
'    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("CaptionAlignment", m_CaptionAlignment, m_def_CaptionAlignment)
    Call PropBag.WriteProperty("Font", picDown.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", picKind.FontBold, 0)
    Call PropBag.WriteProperty("FontSize", picKind.FontSize, 0)
    Call PropBag.WriteProperty("FontName", picKind.FontName, "")
    Call PropBag.WriteProperty("ForeColor", picKind.ForeColor, &H80000012)
    
    Call PropBag.WriteProperty("IDKind", m_IDKind, m_def_IDKind)
    Call PropBag.WriteProperty("ShowPropertySet", m_ShowPropertySet, m_def_ShowPropertySet)
    Call PropBag.WriteProperty("DefaultCardType", m_DefaultCardType, m_def_DefaultCardType)
    Call PropBag.WriteProperty("SmallStyle", m_SmallStyle, m_def_SmallStyle)
    Call PropBag.WriteProperty("KeyShift", m_KeyShift, m_def_KeyShift)
    Call PropBag.WriteProperty("NotAutoAppendKind", m_NotAutoAppendKind, m_def_NotAutoAppendKind)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("objParent", mobjParent, Nothing)
    Call PropBag.WriteProperty("NotContainFastKey", m_NotContainFastKey, m_def_NotContainFastKey)
    Call PropBag.WriteProperty("AllowAutoICCard", m_AllowAutoICCard, m_def_AllowAutoICCard)
    Call PropBag.WriteProperty("AllowAutoIDCard", m_AllowAutoIDCard, m_def_AllowAutoIDCard)
    Call PropBag.WriteProperty("AllowAutoCommCard", m_AllowAutoCommCard, m_def_AllowAutoCommCard)
    Call PropBag.WriteProperty("MustSelectItems", m_MustSelectItems, m_def_MustSelectItems)
      
    Call PropBag.WriteProperty("OnlyReadCardNo", m_OnlyReadCardNo, m_def_OnlyReadCardNo)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HE0E0E0)
    Call PropBag.WriteProperty("SaveRegType", m_SaveRegType, m_def_SaveRegType)
    Call PropBag.WriteProperty("ProductName", m_ProductName, m_def_ProductName)
    Call PropBag.WriteProperty("Locked", m_Locked, m_def_Locked)
End Sub
'注意！不要删除或修改下列被注释的行！
'MemberInfo=1,0,0,0
Public Property Get Appearance() As IDKind_Appearance
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As IDKind_Appearance)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    Call SetCaption
    Call UserControl_Resize
    'Call ReInitCards
    
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=9,0,2,0
Public Property Get Cards() As Object
    Set Cards = m_Cards
End Property

Public Property Set Cards(ByVal New_Cards As Object)
    If Ambient.UserMode = False Then Exit Property
    
    Set m_Cards = New_Cards
    PropertyChanged "Cards"
    If m_Cards Is Nothing Then Exit Property
    
    mstrCardType = ""
    Call InitCardsObject
    Call ReInitCards
    Call FromCardstoIDKindstr
    Call UserControl_Resize
End Property
Private Sub FromCardstoIDKindstr()
    Dim objCard As Card
    Dim strNewIdKindStr As String
    Dim i As Long
    For i = 1 To Cards.Count
        strNewIdKindStr = strNewIdKindStr & ";" & IIf(Cards(i).短名 = "", Left(Cards(i).名称, 1), Cards(i).短名)
        strNewIdKindStr = strNewIdKindStr & "|" & Cards(i).名称
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).是否刷卡, 0, 1)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).接口序号 < 0, 0, Cards(i).接口序号)
        strNewIdKindStr = strNewIdKindStr & "|" & Cards(i).卡号长度
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).缺省标志, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).是否存在帐户, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & Cards(i).卡号密文规则
        '增加是否扫描，是否接触式读卡，是否非接触式读卡性质
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).是否扫描, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).是否接触式读卡, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).是否非接触式读卡, 1, 0)
    Next
    If strNewIdKindStr <> "" Then strNewIdKindStr = Mid(strNewIdKindStr, 2)
    m_IDKindStr = strNewIdKindStr
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,
Public Property Get IDKindStr() As String
    IDKindStr = m_IDKindStr
End Property

Public Property Let IDKindStr(ByVal New_IDKindStr As String)
    m_IDKindStr = New_IDKindStr
    PropertyChanged "IDKindStr"
    Set m_Cards = zlGetKindCards(m_IDKindStr, , NotAutoAppendKind, mOnlyThreeCard)
    If m_Cards.Count = 1 And ShowPropertySet = False Then
        picDown.Visible = False
        mblnSingle = True
    ElseIf m_Cards.Count > 1 Or ShowPropertySet = True Then
        picDown.Visible = True
        mblnSingle = False
    End If
    Call InitCardsObject
    Call ReInitCards
    Call FromCardstoIDKindstr
    Call UserControl_Resize
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14
Public Sub Locale(Optional ByVal intSkip As Integer = 1)
    If intSkip > 0 Then
        '向后
        If m_IDKind + intSkip > mobjCards.Count Then
            m_IDKind = m_IDKind + intSkip - mobjCards.Count
        Else
            m_IDKind = m_IDKind + intSkip
        End If
        If m_IDKind < 1 Or m_IDKind > mobjCards.Count Then
            m_IDKind = mobjCards.Count
        End If
    Else: Print
        '身前
        If m_IDKind + intSkip < 1 Then
            m_IDKind = mobjCards.Count
        Else
            m_IDKind = m_IDKind + intSkip
        End If
    End If
    Err = 0: On Error Resume Next
    mstrCardType = IIf(mobjCards(m_IDKind).接口序号 > 0, mobjCards(m_IDKind).接口序号, mobjCards(m_IDKind).名称)
    Call ReInitCards
    RaiseEvent ItemClick(m_IDKind, mobjCurCard)
'    DoEvents
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As IDKind_BorderStyle
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As IDKind_BorderStyle)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    Call SetCtrlVisible
    Call SetKindWidth
    Call UserControl_Resize
End Property


'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,1
Public Property Get CaptionAlignment() As IDKind_CaptionAlignment
    CaptionAlignment = m_CaptionAlignment
End Property
Public Property Let CaptionAlignment(ByVal New_CaptionAlignment As IDKind_CaptionAlignment)
    m_CaptionAlignment = New_CaptionAlignment
    PropertyChanged "CaptionAlignment"
    Call SetCaption
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=picDown,picDown,-1,Font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set picKind.Font = New_Font
    Set UserControl.Font = New_Font
    Set picDown.Font = New_Font
    PropertyChanged "Font"
    mblnNotItemClick = True '避免触发ItemClick事件
    Call ReInitCards
    Call UserControl_Resize
    mblnNotItemClick = False
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=picKind,picKind,-1,FontBold
Public Property Get FontBold() As Boolean
    FontBold = picKind.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    picKind.FontBold() = New_FontBold
    UserControl.FontBold() = New_FontBold
    PropertyChanged "FontBold"
    mblnNotItemClick = True '避免触发ItemClick事件
    Call ReInitCards
    Call UserControl_Resize
    mblnNotItemClick = False
    
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=picKind,picKind,-1,FontSize
Public Property Get FontSize() As Single
    FontSize = picKind.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    picKind.FontSize() = New_FontSize
    UserControl.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=picKind,picKind,-1,FontName
Public Property Get FontName() As String
    FontName = picKind.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    picKind.FontName() = New_FontName
    UserControl.FontName() = New_FontName
    PropertyChanged "FontName"
    mblnNotItemClick = True '避免触发ItemClick事件
    Call ReInitCards
    Call UserControl_Resize
    mblnNotItemClick = False
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=picKind,picKind,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = picKind.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    picKind.ForeColor() = New_ForeColor
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    mblnNotItemClick = True '避免触发ItemClick事件
    Call ReInitCards
    Call UserControl_Resize
    mblnNotItemClick = False

End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get IDKind() As Integer
    IDKind = m_IDKind
End Property

Public Property Let IDKind(ByVal New_IDKind As Integer)
    Dim objCard As Card
    Set objCard = GetIDKindCard(New_IDKind, CardTypeIndex)
    If objCard Is Nothing Then Exit Property
    m_IDKind = New_IDKind
    PropertyChanged "IDKind"
    Set mobjCurCard = objCard
    mstrCardType = IIf(mobjCurCard.接口序号 > 0, mobjCurCard.接口序号, mobjCurCard.名称)
    mstrPopuCaption = IIf(ShowSortName, mobjCurCard.短名, mobjCurCard.名称)
    picKind.ToolTipText = mobjCurCard.名称
    Call SetCaption
    RaiseEvent ItemClick(m_IDKind, objCard)
'    DoEvents
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14
Public Function GetIDKindCard(ByVal strCardType As String, _
    Optional MachMode As Mach_Mode, Optional bln消费卡 As Boolean = False) As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定卡
    '入参:strCardType-卡类别(数字为指定的卡类别ID,字符为按名称匹配)
    '     MachMode-匹配方式
    '     bln消费卡-是否按消费卡查找
    '返回: 成功，返回卡类别对象;否则返回Nothing
    '编制:刘兴洪
    '日期:2012-08-20 18:20:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, i As Long
    If MachMode <> CardTypeName Then
        lngCardTypeID = Val(strCardType)
        Err = 0: On Error Resume Next
        If MachMode = CardTypeID Then
            For i = 1 To Cards.Count
                If Cards(i).接口序号 = lngCardTypeID And Cards(i).消费卡 = bln消费卡 Then
                    Set GetIDKindCard = Cards(i): Exit Function
                End If
            Next
            Set GetIDKindCard = Nothing
            Exit Function
        Else
            If lngCardTypeID = -1 Then Set GetIDKindCard = Nothing: Exit Function
            Set GetIDKindCard = mobjCards(lngCardTypeID)     '索引只能取有效的卡的索引
        End If
        If Err <> 0 Then Set GetIDKindCard = Nothing
        Exit Function
    End If
    
    For i = 1 To Cards.Count
        Select Case strCardType
        Case "身份证", "身份证号", "二代身份证"
            If InStr(1, Cards(i).名称, "身份证") > 0 Then
                 Set GetIDKindCard = Cards(i): Exit Function
            End If
        Case Else
            If strCardType = Cards(i).名称 Then
                 Set GetIDKindCard = Cards(i): Exit Function
            End If
        End Select
    Next
    Set GetIDKindCard = Nothing
End Function


'注意！不要删除或修改下列被注释的行！
'MemberInfo=14
Public Function GetKindIndex(ByVal strCardType As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定卡
    '入参:strCardType-卡类别(数字为指定的卡类别ID,字符为按名称匹配)
    '返回: 成功，返回索引值,否则返回-1
    '编制:刘兴洪
    '日期:2012-08-20 18:20:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, i As Long
    Dim blnCardTypeID As Boolean
    
    blnCardTypeID = IsNumeric(strCardType)
    lngCardTypeID = Val(strCardType)
    For i = 1 To mobjCards.Count
        If blnCardTypeID Then
            If mobjCards(i).接口序号 = lngCardTypeID Then GetKindIndex = i: Exit Function
        Else
            Select Case strCardType
            Case "身份证", "身份证号", "二代身份证"
                If InStr(1, mobjCards(i).名称, "身份证") > 0 Then
                     GetKindIndex = i: Exit Function
                End If
            Case "IC卡", "IC卡号"
                If InStr(1, mobjCards(i).名称, "IC卡") > 0 Then
                     GetKindIndex = i: Exit Function
                End If
                
            Case Else
                If strCardType Like "姓名*" And mobjCards(i).名称 Like "姓名*" Then
                         GetKindIndex = i: Exit Function
                Else
                    If strCardType = mobjCards(i).名称 Then
                         GetKindIndex = i: Exit Function
                    End If
                End If
            End Select
        End If
    Next
    GetKindIndex = -1
End Function
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get ShowPropertySet() As Boolean
    ShowPropertySet = m_ShowPropertySet
End Property

Public Property Let ShowPropertySet(ByVal New_ShowPropertySet As Boolean)
    m_ShowPropertySet = New_ShowPropertySet
    PropertyChanged "ShowPropertySet"
    
End Property
Public Property Get GetCurCard() As Object
    Set GetCurCard = mobjCurCard
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,
Public Property Get DefaultCardType() As String
    DefaultCardType = m_DefaultCardType
End Property

Public Property Let DefaultCardType(ByVal New_DefaultCardType As String)
    m_DefaultCardType = New_DefaultCardType
    PropertyChanged "DefaultCardType"
    
    If m_DefaultCardType = "" Then
        Set mobjDefaultCard = Nothing
    ElseIf IsNumeric(m_DefaultCardType) Then
        Set mobjDefaultCard = GetIDKindCard(m_DefaultCardType, CardTypeID)
    Else
        Set mobjDefaultCard = GetIDKindCard(m_DefaultCardType, CardTypeName)
    End If
End Property
Public Property Get GetfaultCard() As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省卡
    '编制:刘兴洪
    '日期:2012-08-24 10:26:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If DefaultCardType <> "" And mobjDefaultCard Is Nothing Then
        If IsNumeric(DefaultCardType) Then
            Set mobjDefaultCard = GetIDKindCard(DefaultCardType, CardTypeID)
        Else
            Set mobjDefaultCard = GetIDKindCard(DefaultCardType, CardTypeName)
        End If
    End If
    Set GetfaultCard = mobjDefaultCard
End Property


'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get SmallStyle() As Boolean
    SmallStyle = m_SmallStyle
End Property

Public Property Let SmallStyle(ByVal New_SmallStyle As Boolean)
    Dim MyFont As New StdFont
    '为了与老的Kind歉容
    m_SmallStyle = New_SmallStyle
    PropertyChanged "SmallStyle"
    If New_SmallStyle Then
        MyFont.Size = 10
    Else
        MyFont.Size = 12
    End If
    Set UserControl.Font = MyFont
    Set picKind.Font = MyFont
    Call SetCaption
    
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,2
Public Property Get KeyShift() As Long
    KeyShift = m_KeyShift
End Property

Public Property Let KeyShift(ByVal New_KeyShift As Long)
    If New_KeyShift < 256 Then
        m_KeyShift = New_KeyShift
        PropertyChanged "KeyShift"
    Else
        MsgBox "无效的属性值,请参考KeyDown事件的shift参数", vbInformation, App.ProductName
    End If
End Property

Private Sub mobjParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If UserControl.Enabled Then
        If Shift = KeyShift Then
            'keycode:96是小键盘的0,105是9
            If KeyCode > 95 And KeyCode < 106 Then
                IDKind = KeyCode - 96
            ElseIf KeyCode = 123 Then   'Ctrol+F12,对允许执行点击操作的都可以响应该热键
                Call picKind_MouseDown(1, Shift, 0, 0)
            End If
        End If
    End If
End Sub
'注意！不要删除或修改下列被注释的行！
'MemberInfo=7
Public Function ListCount() As Integer
    If mobjCards Is Nothing Then Exit Function
    ListCount = mobjCards.Count
End Function

Public Function MacthKey(strFunKey As String, strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 匹配按钮是否正确
    '入参:KeyCode-按键
    '       Shift-功能键
    '返回: 匹配成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-08-23 09:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intFun As Integer
    Dim kbState(255) As Byte, lngReturn As Long
    Dim blnDownCTRL As Boolean
    Dim blnSHIFT As Boolean
    Dim blnFun As Boolean '功能键是否满足
    Dim blnTemp As Boolean
    
    lngReturn = GetKeyboardState(kbState(0))
    If lngReturn = 0 Then Exit Function
    
    blnDownCTRL = kbState(vbKeyControl) And &H80
    blnSHIFT = kbState(vbKeyShift) And &H80
    
    strFunKey = UCase(strFunKey)
    strKey = UCase(strKey)
    
    blnFun = IIf(strFunKey = "CTRL", blnDownCTRL, IIf(strFunKey = "SHIFT", blnSHIFT, Not (blnDownCTRL Or blnSHIFT)))
    
    If Left(strKey, 1) = "F" And Len(strKey) >= 2 Then
        'F1-Fn打头
        i = vbKeyF1 - 1 + Val(Mid(strKey, 2))
        MacthKey = kbState(i) And &H80 And blnFun
 
        Exit Function
    End If
 
    If strKey = " " Or strKey = "空格键" Then
        MacthKey = kbState(vbKeySpace) And &H80 And blnFun
        Exit Function
    End If
         
 
    '上下左右
    If InStr(" ←↑→↓", strKey) > 0 Then
        i = Switch(strKey = "←", vbKeyLeft, strKey = "↑", vbKeyUp, strKey = "→", vbKeyRight, True, vbKeyDown)
        MacthKey = kbState(i) And &H80 And blnFun
        Exit Function
    End If
   
    i = Asc(strKey)
    If i >= Asc("A") And i <= Asc("Z") Then
        '字母
        i = i - Asc("A")
        i = vbKeyA + i
        MacthKey = kbState(i) And &H80 And blnFun
        Exit Function
    End If
    
    '小键盘的键(数字键盘)
    If Left(strKey, 3) = "NUM" And Len(strKey) > 3 Then
        strKey = Mid(strKey, 4)
        If Val(strKey) >= 0 And Val(strKey) <= 9 And IsNumeric(strKey) Then
            i = vbKeyNumpad0 + Val(strKey)
            MacthKey = kbState(i) And &H80 And blnFun
            Exit Function
        End If
        Select Case strKey
        Case "*"
            MacthKey = kbState(vbKeyMultiply) And &H80 And blnFun
        Case "+"
            MacthKey = kbState(vbKeyAdd) And &H80 And blnFun
        Case "-"
            MacthKey = kbState(vbKeySubtract) And &H80 And blnFun
        Case "/"
            MacthKey = kbState(vbKeyDivide) And &H80 And blnFun
        Case "."
            MacthKey = kbState(vbKeyDecimal) And &H80 And blnFun
        Case "ENTER"
            MacthKey = kbState(vbKeySeparator) And &H80 And blnFun
        End Select
        Exit Function
    End If
    
    Select Case strKey
    Case ":"
        blnTemp = kbState(186) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case ";"
        blnTemp = kbState(186) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case """"
        blnTemp = kbState(222) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "'"
        blnTemp = kbState(222) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "?"
        blnTemp = kbState(191) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "/"
        blnTemp = kbState(191) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "|"
        blnTemp = kbState(220) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "\"
        blnTemp = kbState(220) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "{"
        blnTemp = kbState(219) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "["
        blnTemp = kbState(219) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "}"
        blnTemp = kbState(221) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "]"
        blnTemp = kbState(221) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "~"
        blnTemp = kbState(192) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "`"
        blnTemp = kbState(192) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "<"
        blnTemp = kbState(188) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case ","
        blnTemp = kbState(188) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case ">"
        blnTemp = kbState(190) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "."
        blnTemp = kbState(190) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "+"
        blnTemp = kbState(187) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "="
        blnTemp = kbState(187) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "_"
        blnTemp = kbState(189) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "-"
        blnTemp = kbState(189) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "!"
        blnTemp = kbState(vbKey1) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "1"
        blnTemp = kbState(vbKey1) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "@"
        blnTemp = kbState(vbKey2) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "2"
        blnTemp = kbState(vbKey2) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "#"
        blnTemp = kbState(vbKey3) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "3"
        blnTemp = kbState(vbKey3) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "$"
        blnTemp = kbState(vbKey4) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "4"
        blnTemp = kbState(vbKey4) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "%"
        blnTemp = kbState(vbKey5) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "5"
        blnTemp = kbState(vbKey5) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "^"
        blnTemp = kbState(vbKey6) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "6"
        blnTemp = kbState(vbKey6) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "&"
        blnTemp = kbState(vbKey7) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "7"
        blnTemp = kbState(vbKey7) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "*"
        blnTemp = kbState(vbKey8) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "8"
        blnTemp = kbState(vbKey8) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case "("
        blnTemp = kbState(vbKey9) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "9"
        blnTemp = kbState(vbKey9) And &H80
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case ")"
        blnTemp = kbState(vbKey0) And &H80 And blnSHIFT: MacthKey = blnTemp And blnFun: Exit Function
    Case "0"
        blnTemp = kbState(vbKey0) And &H80:
        blnTemp = blnTemp And IIf(strFunKey <> "SHIFT", Not blnSHIFT, True)
        MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyLButton")    'vbKeyLButton 0x1 鼠标左键
        blnTemp = kbState(vbKeyLButton) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyMButton")   'vbKeyRButton 0x2 鼠标右键'
        blnTemp = kbState(vbKeyRButton) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyMButton")   'vbKeyMButton 0x4 鼠标中键
        blnTemp = kbState(vbKeyMButton) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyCancel")   ' vbKeyCancel 0x3 CANCEL 键
        blnTemp = kbState(vbKeyCancel) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyBack")   ' vbKeyBack 0x8 BACKSPACE 键
        blnTemp = kbState(vbKeyBack) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyTab")   ' vbKeyTab 0x9 TAB 键
        blnTemp = kbState(vbKeyTab) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyClear")   ' vbKeyClear 0xC CLEAR 键
        blnTemp = kbState(vbKeyClear) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("ENTER")   ' vbKeyReturn 0xD ENTER 键
        blnTemp = kbState(vbKeyReturn) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("SHIFT")   ' vbKeyShift 0x10 SHIFT 键
        blnTemp = kbState(vbKeyShift) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("CTRL")   ' vbKeyControl 0x11 CTRL 键
        blnTemp = kbState(vbKeyControl) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyMenu")   ' vbKeyMenu 0x12 MENU 键
        blnTemp = kbState(vbKeyMenu) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyPause")   ' vbKeyPause 0x13 PAUSE 键
        blnTemp = kbState(vbKeyPause) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase(" CAPS LOCK")   ' vbKeyCapital 0x14 CAPS LOCK 键
        blnTemp = kbState(vbKeyCapital) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("ESC")   ' vbKeyEscape 0x1B ESC 键
        blnTemp = kbState(vbKeyEscape) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("Space")   ' vbKeySpace 0x20 SPACEBAR 键
        blnTemp = kbState(vbKeySpace) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("SELECT")   ' vbKeySelect 0x29 SELECT 键
        blnTemp = kbState(vbKeySelect) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("PRINT SCREEN")   ' vbKeyPrint 0x2A PRINT SCREEN 键
        blnTemp = kbState(vbKeyPrint) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("EXECUTE")   ' vbKeyExecute 0x2B EXECUTE 键
        blnTemp = kbState(vbKeyPrint) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("SNAPSHOT")   ' vbKeySnapshot 0x2C SNAPSHOT 键
        blnTemp = kbState(vbKeySnapshot) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("INSERT")   ' vbKeyInsert 0x2D INSERT 键
        blnTemp = kbState(vbKeyInsert) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("DELETE")   ' vbKeyDelete 0x2E DELETE 键
        blnTemp = kbState(vbKeyDelete) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("HELP")   ' vbKeyHelp 0x2F HELP 键
        blnTemp = kbState(vbKeyHelp) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("NUM LOCK")   ' vbKeyNumlock 0x90 NUM LOCK 键
        blnTemp = kbState(vbKeyNumlock) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case Else
    End Select
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    picDown.Enabled = New_Enabled
    picDown.Enabled = New_Enabled
    
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    mblnNotItemClick = True '避免触发ItemClick事件
    Call ReInitCards
    Call UserControl_Resize
    mblnNotItemClick = False

End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=9,0,0,0
Public Property Get objParent() As Object
    Set objParent = mobjParent
End Property
Public Property Set objParent(ByVal New_objParent As Object)
    Set mobjParent = New_objParent
    PropertyChanged "objParent"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,F1;CTRL+F1;F12;CTRL+F12
Public Property Get NotContainFastKey() As String
    NotContainFastKey = m_NotContainFastKey
End Property

Public Property Let NotContainFastKey(ByVal New_NotContainFastKey As String)
    m_NotContainFastKey = New_NotContainFastKey
    PropertyChanged "NotContainFastKey"
    
End Property
Public Sub zlInit(ByVal frmMain As Object, Optional ByVal lngSys As Long, Optional ByVal lngModul As Long, _
    Optional cnOracle As ADODB.Connection, Optional strDBUser As String, Optional objPublicOneCard As Object, _
    Optional strIDKindStr As String = "", Optional objBoundTextBox As Object, Optional strProductName As String = "", _
    Optional blnOnlyThreeCard As Boolean = False, Optional ByVal blnIsObjRegisterAlone As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化
    '入参:frmMain-调用的主窗口
    '     lngSys : 系统编号
    '     lngModul:需要执行的功能序号
    '     objCardSquare-卡结算部件
    '     cnOracle:主程序的数据库连接
    '     strIDKindStr-身份识别的类别项,有两种格式:
    '               一种是缺省的:短名1|全名1|读卡标志1;…. ;短名n|全名n|读卡标志n
    '               一种是扩展格式:短名1|全名1|读卡标志1|卡类别ID1|卡号长度1|缺省标志1(1-当前缺省;0-非缺省)|是否存在帐户1(1-存在帐户;0-不存在帐户)|卡号密文1(第几位至第几位加密,空为不加密);…
    '   strProductName-产品模块名称(主要用于相关模块参数的保存)
    '    blnOnlyThreeCard-仅三方卡
    '    blnIsObjRegisterAlone-是否使用独立的注册部件(True:使用:zlRegisterAlone.DLL,否则使用zlRegister.dll)
    '编制:刘兴洪
    '日期:2012-08-16 10:45:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Set gobjPubOneCard = objPublicOneCard: gstrDBUser = strDBUser
    glngModul = lngModul: glngSys = lngSys:
    Set mcnOracle = cnOracle: Set gcnOracle = cnOracle
    
    gblnIsObjRegisterAlone = blnIsObjRegisterAlone
    
    Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    
    mOnlyThreeCard = blnOnlyThreeCard
    If strProductName <> "" Then m_ProductName = strProductName
    If m_ProductName = "" Then m_ProductName = "IDKindNew"
    
    Set mobjParent = frmMain
    If TypeName(objBoundTextBox) = "TextBox" Then Set mobjTxtInput = objBoundTextBox
    If gcnOracle Is Nothing Then GoTo GoInitKind
    If gcnOracle.State <> 1 Then GoTo GoInitKind
    Call zlInitCards(mcnOracle, mRegType)
GoInitKind:
    If strIDKindStr = "" Then
        strIDKindStr = IDKindStr
    End If
    Set Cards = zlGetKindCards(strIDKindStr, False, NotAutoAppendKind, mOnlyThreeCard)
    If Cards.Count = 1 And ShowPropertySet = False Then
        picDown.Visible = False
        mblnSingle = True
        Call UserControl_Resize
    End If
    '创建公共对象
    Call CreateComEvtsObject
    '80115:李南春,2014/12/1,初始化当前卡类别
    mstrCardType = "": Call ReInitCards
    
End Sub

Private Sub CreateComEvtsObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建公共事件对象
    '返回: 创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-08-28 16:16:00
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If mobjPubOneCard Is Nothing Then Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    '创建公共对象
    Err = 0: On Error Resume Next
    
    If mobjICCard Is Nothing And AllowAutoICCard Then
        Set mobjICCard = New clsICCard
        If mobjParent Is Nothing Then
            Call mobjICCard.SetParent(UserControl.hWnd)
        Else
            Call mobjICCard.SetParent(mobjParent.hWnd)
        End If
         Set mobjICCard.gcnOracle = mcnOracle
    End If
    If mobjCommEvents Is Nothing And AllowAutoCommCard And Not mobjPubOneCard Is Nothing Then
        Set mobjCommEvents = New clsCommEvents
        If Not mobjParent Is Nothing Then
             Call mobjPubOneCard.objThirdSwap.zlInitEvents(mobjParent.hWnd, mobjCommEvents)
        Else
             Call mobjPubOneCard.objThirdSwap.zlInitEvents(UserControl.hWnd, mobjCommEvents)
        End If
    End If
    
    If mobjIDCard Is Nothing And AllowAutoIDCard Then
        Set mobjIDCard = New clsIDCard
        If mobjParent Is Nothing Then
            Call mobjIDCard.SetParent(UserControl.hWnd)
        Else
            Call mobjIDCard.SetParent(mobjParent.hWnd)
        End If
    End If
    
End Sub
Private Sub CloseComEvntsObject()
    '关闭相关对象
    Err = 0: On Error Resume Next
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
    End If

    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
    End If
    Set mobjCommEvents = Nothing
    Set mobjIDCard = Nothing
    Set mobjICCard = Nothing
    mOnlyThreeCard = False
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14
Public Sub SetAutoReadCard(ByVal blnAutoReadCard As Boolean)
    '设置自动读卡
    If AllowAutoICCard And Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (blnAutoReadCard)
    If AllowAutoIDCard And Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (blnAutoReadCard)
    If Not AllowAutoCommCard Then Exit Sub
    
    If mobjPubOneCard Is Nothing Then Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    
    If mobjPubOneCard Is Nothing Then Exit Sub
    If Not mobjParent Is Nothing Then
         Call mobjPubOneCard.objThirdSwap.zlInitEvents(mobjParent.hWnd, mobjCommEvents)
    Else
         Call mobjPubOneCard.objThirdSwap.zlInitEvents(UserControl.hWnd, mobjCommEvents)
    End If
    Err = 0: On Error Resume Next
    Call mobjPubOneCard.objThirdSwap.SetEnabled(blnAutoReadCard)
    If Err <> 0 Then Err = 0: On Error GoTo 0
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get AllowAutoICCard() As Boolean
    AllowAutoICCard = m_AllowAutoICCard
End Property

Public Property Let AllowAutoICCard(ByVal New_AllowAutoICCard As Boolean)
    m_AllowAutoICCard = New_AllowAutoICCard
    PropertyChanged "AllowAutoICCard"
    '76256,冉俊明,2014-8-5,先初始化IDKind(zlInit),后设置该属性,未创建公共事件对象
    '创建公共对象
    Call CreateComEvtsObject
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get AllowAutoIDCard() As Boolean
    AllowAutoIDCard = m_AllowAutoIDCard
End Property

Public Property Let AllowAutoIDCard(ByVal New_AllowAutoIDCard As Boolean)
    m_AllowAutoIDCard = New_AllowAutoIDCard
    PropertyChanged "AllowAutoIDCard"
    '76256,冉俊明,2014-8-5,先初始化IDKind(zlInit),后设置该属性,未创建公共事件对象
    '创建公共对象
    Call CreateComEvtsObject
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get AllowAutoCommCard() As Boolean
    AllowAutoCommCard = m_AllowAutoCommCard
End Property

Public Property Let AllowAutoCommCard(ByVal New_AllowAutoCommCard As Boolean)
    m_AllowAutoCommCard = New_AllowAutoCommCard
    PropertyChanged "AllowAutoCommCard"
    '76256,冉俊明,2014-8-5,先初始化IDKind(zlInit),后设置该属性,未创建公共事件对象
    '创建公共对象
    Call CreateComEvtsObject
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get NotAutoAppendKind() As Boolean
    NotAutoAppendKind = m_NotAutoAppendKind
End Property

Public Property Let NotAutoAppendKind(ByVal New_NotAutoAppendKind As Boolean)
    m_NotAutoAppendKind = New_NotAutoAppendKind
    PropertyChanged "NotAutoAppendKind"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,""
Public Property Get MustSelectItems() As String
    MustSelectItems = m_MustSelectItems
End Property

Public Property Let MustSelectItems(ByVal New_MustSelectItems As String)
    m_MustSelectItems = New_MustSelectItems
    PropertyChanged "MustSelectItems"
End Property

Public Sub RaisEffect(picBox As Object, Optional intStyle As Integer, _
    Optional strName As String = "", Optional TxtAlignment As Integer)
    Select Case TxtAlignment
    Case 0
        Call zlRaisEffect(picBox, intStyle, strName, mLeftAgnmt)
    Case 1
        Call zlRaisEffect(picBox, intStyle, strName, mRightAgnmt)
    Case 2
        Call zlRaisEffect(picBox, intStyle, strName, mRightAgnmt)
    End Select
End Sub

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,1,0,False
Public Property Get DefaultCardShowPassText() As Boolean
    If m_Cards.按缺省卡查找 And Not mobjDefaultCard Is Nothing Then
        DefaultCardShowPassText = mobjDefaultCard.卡号密文规则 <> ""
    Else
        If mobjCards Is Nothing Then
             DefaultCardShowPassText = Cards.加密显示
        Else
             DefaultCardShowPassText = mobjCards.加密显示
        End If
    End If
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,1,0,False
Public Property Get ShowPassText() As Boolean
    If mobjCurCard Is Nothing Then Exit Property
    If mobjCurCard.模糊查找项 Then
        ShowPassText = DefaultCardShowPassText
    Else
         ShowPassText = IIf(mobjCurCard.卡号密文规则 = "0", "", mobjCurCard.卡号密文规则) <> ""
    End If
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,1,0,100
Public Property Get GetDefaultCardNoLen() As Integer
    If m_Cards Is Nothing Then GetDefaultCardNoLen = 100: Exit Property
    If m_Cards.按缺省卡查找 And Not mobjDefaultCard Is Nothing Then
        GetDefaultCardNoLen = mobjDefaultCard.卡号长度
    Else
        GetDefaultCardNoLen = 100
    End If
End Property
Public Property Get GetCardNoLen() As Integer
    If mobjCurCard Is Nothing Then Exit Property
    If mobjCurCard.模糊查找项 Then
        GetCardNoLen = GetDefaultCardNoLen
    Else
         GetCardNoLen = mobjCurCard.卡号长度
    End If
End Property

Private Function ActiveFastKeyInside() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:激活快键
    '返回: 快键激活后,返回True,否则返回False
    '编制:刘兴洪
    '日期:2012-08-29 22:58:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '激活快键
    Dim objCurCard As Card, bln缺省读卡 As Boolean
    bln缺省读卡 = CheckReadCard
    If mobjCurCard Is Nothing Or m_Cards Is Nothing Then Exit Function
    '快键操作IDKind
     If m_Cards.向前滚动功能键 <> "" Or m_Cards.向前滚动快键 <> "" Then
            If MacthKey(m_Cards.向前滚动功能键, m_Cards.向前滚动快键) Then
                Call Locale(-1): ActiveFastKeyInside = True: Exit Function
            End If
    Else
        '缺省
        If MacthKey("SHIFT", "F4") Then Locale (-1): ActiveFastKeyInside = True: Exit Function
        '缺省
        If MacthKey("", "↑") Then Locale (-1): ActiveFastKeyInside = True: Exit Function
        If MacthKey("", "PgUp") Then Locale (-1): ActiveFastKeyInside = True: Exit Function
     End If
     If m_Cards.向后滚动功能键 <> "" Or m_Cards.向后滚动快键 <> "" Then
            If MacthKey(m_Cards.向后滚动功能键, m_Cards.向后滚动快键) Then
                Call Locale(1): ActiveFastKeyInside = True: Exit Function
            End If
    Else
        '缺省
        If MacthKey("", "F4") Then Locale (1): ActiveFastKeyInside = True: Exit Function
        '缺省
        If MacthKey("", "↓") Then Locale (1): ActiveFastKeyInside = True: Exit Function
        If MacthKey("", "PgDn") Then Locale (-1): ActiveFastKeyInside = True: Exit Function
     End If
    '读卡
    If mobjCurCard.是否接触式读卡 = False And bln缺省读卡 = False Then Exit Function
    If MacthKey(m_Cards.读卡功能键, m_Cards.读卡快键) = False Then Exit Function
    If bln缺省读卡 Then
        Set objCurCard = mobjCurCard
        Set mobjCurCard = Cards("K" & mlngDefaultCardID)
    End If
    Call IDKindClick(mobjCurCard, True)
    If bln缺省读卡 Then Set mobjCurCard = objCurCard
    ActiveFastKeyInside = True
End Function

Public Function ActiveFastKey() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:激活快键
    '返回: 快键激活后,返回True,否则返回False
    '编制:刘兴洪
    '日期:2012-08-29 22:58:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnTextFocus = False Then
        ActiveFastKey = ActiveFastKeyInside
    End If
End Function
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,True
Public Property Get OnlyReadCardNo() As Boolean
    OnlyReadCardNo = m_OnlyReadCardNo
End Property
Public Property Let OnlyReadCardNo(ByVal New_OnlyReadCardNo As Boolean)
    m_OnlyReadCardNo = New_OnlyReadCardNo
    PropertyChanged "OnlyReadCardNo"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    picDown.BackColor = New_BackColor
    picKind.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property
Public Sub Refrash()
    Call RePicKindStatu
End Sub
Public Function GetDefaultCardTypeID() As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取缺省的卡类别ID
    '返回: 缺省的卡类别ID;-1表示按模糊查找进行查找病人
    '编制:刘兴洪
    '日期:2012-08-31 17:02:08
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng卡类别ID As Long
    If m_Cards.按缺省卡查找 And Not mobjDefaultCard Is Nothing Then
        lng卡类别ID = mobjDefaultCard.接口序号
    Else
        lng卡类别ID = "-1"
    End If
    '无意中删除了,重新补上
    GetDefaultCardTypeID = lng卡类别ID
End Function
'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,3
Public Property Get SaveRegType() As IDKind_RegType
    SaveRegType = m_SaveRegType
End Property

Public Property Let SaveRegType(ByVal New_SaveRegType As IDKind_RegType)
    m_SaveRegType = New_SaveRegType
    Call FromRegType(m_SaveRegType)
    PropertyChanged "SaveRegType"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,
Public Property Get ProductName() As String
    ProductName = m_ProductName
End Property

Public Property Let ProductName(ByVal New_ProductName As String)
    m_ProductName = New_ProductName
    gstrProductName = m_ProductName
    PropertyChanged "ProductName"
End Property

Public Property Get AvailabilityIdkindStr() As String
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效的IDKindStr
    '返回:有效的IDkindStr:短名|全名|读卡性质|接口序号|卡号长度|缺省标志|是否存在帐户|卡号密文规则;...
    '编制:刘兴洪
    '日期:2012-10-22 11:55:31
    '说明:(目前主要应用于自动测试)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim strNewIdKindStr As String
    Dim i As Long
    
    If mobjCards Is Nothing Then Exit Function
    strNewIdKindStr = ""
    For i = 1 To mobjCards.Count
        Set objCard = mobjCards(i)
        strNewIdKindStr = strNewIdKindStr & ";" & IIf(objCard.短名 = "", Left(objCard.名称, 1), objCard.短名)
        strNewIdKindStr = strNewIdKindStr & "|" & objCard.名称
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).是否刷卡, 0, 1)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(objCard.接口序号 < 0, 0, objCard.接口序号)
        strNewIdKindStr = strNewIdKindStr & "|" & objCard.卡号长度
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(objCard.缺省标志, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(objCard.是否存在帐户, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & objCard.卡号密文规则
        '增加是否扫描，是否接触式读卡，是否非接触式读卡性质
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).是否扫描, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).是否接触式读卡, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).是否非接触式读卡, 1, 0)
    Next
    If strNewIdKindStr <> "" Then strNewIdKindStr = Mid(strNewIdKindStr, 2)
    AvailabilityIdkindStr = strNewIdKindStr
End Property
 
Public Property Get AvailabilityCards() As Object
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取有效的Cards对象
    '返回:有效的Cards对象
    '编制:刘兴洪
    '日期:2012-10-22 11:55:31
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set AvailabilityCards = mobjCards
End Property
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Function CheckReadCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否存在缺省的读卡类别
    '返回: true-存在且有效，false-不存在
    '编制:李南春
    '日期:2014/11/26
    '说明:
    '问题:78768
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCurCard As Card
    If (mobjCurCard.名称 = "姓名" Or mobjCurCard.名称 = "姓名或就诊卡") And mlngDefaultCardID > 0 Then
        Err = 0: On Error Resume Next
        Set objCurCard = Cards("K" & mlngDefaultCardID)
        If Err <> 0 Then Err = 0: Exit Function
        If Not (objCurCard.是否接触式读卡 Or objCurCard.是否非接触式读卡) Then Exit Function
        CheckReadCard = True
    End If
End Function
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get Locked() As Boolean
    Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    PropertyChanged "Locked"
End Property

Public Function SetBrushCardObject(ByVal blnComm As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置刷卡接口
    '返回: true-成功，false-失败
    '编制:李南春
    '日期:2015/7/23 10:37:39
    '问题:85565
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    Dim objCard As Card
    
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    '118959:李南春,2018/1/3,如果支持射频卡，则传入事件对象
    If (mobjCurCard Is Nothing And mobjDefaultCard Is Nothing) Or mobjTxtInput Is Nothing Then Exit Function
    Set objCard = mobjCurCard
    If mobjCurCard.接口序号 = 0 And mobjCurCard.名称 Like "*姓名*" Then Set objCard = mobjDefaultCard

    If objCard.接口序号 = 0 Or objCard.接口程序名 = "" Or Not (objCard.是否刷卡 Or objCard.是否扫描) Then Exit Function
    
    If mobjPubOneCard Is Nothing Then Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    If mobjPubOneCard Is Nothing Then Exit Function
    
    If mobjPubOneCard.objThirdSwap.zlSetBrushCardObject(objCard.接口序号, IIf(blnComm, mobjTxtInput, Nothing), strExpend, objCard.消费卡) Then
        If mobjCommEvents Is Nothing And AllowAutoCommCard = True Then
            Set mobjCommEvents = New clsCommEvents
            Call mobjPubOneCard.objThirdSwap.zlInitEvents(UserControl.hWnd, mobjCommEvents)
        End If
    End If
End Function

Public Function zlFindPatient(ByVal strInput As String, objPatiInfor As Object, _
    Optional objCard As Object, Optional strErrMsg As String, Optional ByVal blnBrushCard As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:通过输入的值或刷卡或读卡的值查找病人并返回病人信息
    '入参:strInput-输入的值或刷卡或读卡的值
    '     objCard-当前的卡类别，nothing时，为IDKindNew. CurCard，否则以传入为准
    '     blnBrushCard-是否是刷卡
    '出参: objPati-当前读卡的病人信息
    '      strErrMsg-返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-09-24 15:13:43
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set objPatiInfor = New clsPatiInfor
    If strInput = "" And mobjTxtInput Is Nothing Then Exit Function
    If strInput = "" And Not mobjTxtInput Is Nothing Then strInput = Trim(mobjTxtInput.Text)
    If strInput = "" Then Exit Function
    zlFindPatient = GetPatient(strInput, objPatiInfor, objCard, strErrMsg)
End Function

Public Function zlGetPatiInforFromPatiID(ByVal lng病人ID As Long, ByRef objPatiInfor As Object, Optional ByRef strErrMsg As String, Optional strOtherName As String = "", Optional strOtherValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:通过病人ID获取病人信息
    '入参:lng病人ID
    '出参: objPati-当前读卡的病人信息
    '      strErrMsg-返回错误信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-09-24 15:13:43
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set objPatiInfor = New clsPatiInfor
    If lng病人ID = 0 Then Exit Function
    zlGetPatiInforFromPatiID = GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPatiInfor, strErrMsg, strOtherName, strOtherValue)
End Function

Public Function zlGetPatiIDFromCardType(ByVal strCardType As String, ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef lng病人ID As Long, _
    Optional ByRef strCardPassWord As String, Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, Optional objCtl As Object = Nothing, Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional ByRef blnCertificate As Boolean = False, Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0, Optional ByVal blnNotCheckValidDate As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的医疗类别和卡号,获取对应的病人ID
    '入参:strCardType-卡类别,如果为数字,这为卡类别ID,如果为字符,则为类别名称
    '       strCardNo-卡号
    '       blnNotShowErrMsg-不显示错误的提示信息
    '       frmMain-调用的主窗体
    '       objCtl-调用的控件
    '       blnShowMergePati-当出现多个满足条件的病人时,是否显示合并功能按钮
    '       blnOnlyContractPati-签约病人
    '       blnUserCancel-选择器中，用户选择了取消
    '       lngShowCardNoTypeID-过滤出多条病信息时，弹出选择器中显示的卡号的卡类别ID,0-表示不显示卡号；>0表示显示指定卡号类别的ID
    '       blnNotCheckValidDate-是否对卡终止使用时间进行检查,true-不检查终止使用时间,false-检查
    '出参:strErrMsg-返回的错误信息
    '       lng病人ID-返回的病人ID
    '       strCardPass-返回卡号的密码
    '       lngCardTypeID-返回卡类别ID(0表示不能确定卡类别ID)
    '返回:获取病人ID成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-14 17:07:51
    '说明:只有存在医疗类别的才调用此函数
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiIDFromCardType = GetPatiIDFromCardType(mcnOracle, strCardType, strCardNo, blnNotShowErrMsg, lng病人ID, _
        strCardPassWord, strErrMsg, lngCardTypeID, objCtl, frmMain, blnShowMergePati, blnOnlyContractPati, _
        blnCertificate, blnUserCancel, lngShowCardNoTypeID, blnNotCheckValidDate)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetPatient(ByVal strInput As String, objPati As clsPatiInfor, ByRef objCard As Card, _
    Optional ByRef strErrMsg As String, Optional ByVal blnBrushCard As Boolean, _
    Optional ByVal lngDefaultCardTypeID As Long, Optional ByVal bln缺省卡查找 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人ID
    '入参:strInput-输入的相关值
    '     objCard-当前读取的卡类别(可以为出参)
    '     blnBrushCard-是否是刷卡
    '出参:lng病人ID-返回:病人ID
    '     objCard-返回读取的卡类别
    '     strErrMsg-返回的错误信息
    '返回: 成功,返回指定的病人ID
    '编制:刘兴洪
    '日期:2012-08-20 17:34:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardPassWord As String, lngCardTypeID As Long
    Dim lng病人ID As Long
    
    On Error GoTo errHandle
    
    '指定卡类别
    If objCard Is Nothing Then Set objCard = GetCurCard
    If objCard.接口序号 > 0 Then
        If GetPatiIDFromCardType(mcnOracle, objCard.接口序号, strInput, False, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID) Then
            If lngCardTypeID <> objCard.接口序号 And lngCardTypeID <> 0 Then
                 Set objCard = GetIDKindCard(lngCardTypeID, CardTypeID)
            End If
        ElseIf IsMobileNo(strInput) Then
            '103000：李南春，2017/2/7，按手机号查找
            If GetPatiIDFromCardType(mcnOracle, "手机号", strInput, False, lng病人ID, strCardPassWord, strErrMsg) Then
                Set objCard = GetIDKindCard("手机号", CardTypeName)
            Else
                lng病人ID = 0
            End If
        Else
            lng病人ID = 0
        End If
        If lng病人ID = 0 Then GoTo NotFindPati:
        If GetPatiIDFromCardType(mcnOracle, lng病人ID, objPati, strErrMsg) = False Then Exit Function
        objPati.卡号 = strInput: objPati.密码 = strCardPassWord
        GetPatient = True: Exit Function
    End If
    If objCard.名称 Like "姓名*" Or objCard.模糊查找项 Then
        If bln缺省卡查找 Then
            lngCardTypeID = lngDefaultCardTypeID
        Else
            lngCardTypeID = -1
        End If
        If blnBrushCard Then
            If GetPatiIDFromCardType(mcnOracle, lngCardTypeID, strInput, False, lng病人ID, strCardPassWord, strErrMsg, _
                lngCardTypeID) = False Then lng病人ID = 0
            If lngCardTypeID <> objCard.接口序号 And lngCardTypeID <> 0 Then
                 Set objCard = GetIDKindCard(lngCardTypeID, CardTypeID)
            ElseIf Not GetfaultCard Is Nothing Then
                Set objCard = GetfaultCard
            End If
            If lng病人ID = 0 Then GoTo NotFindPati:
            If GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg) = False Then Exit Function
            objPati.卡号 = strInput: objPati.密码 = strCardPassWord
            GetPatient = True: Exit Function
        End If
    End If
    Select Case objCard.名称
    Case "姓名", "姓名或就诊卡"
        If Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
            lng病人ID = Val(Mid(strInput, 2))
            If GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg) = False Then Exit Function
        ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
            If GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "门诊号", Mid(strInput, 2)) = False Then Exit Function
        ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
            If GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "住院号", Mid(strInput, 2)) = False Then Exit Function
        ElseIf IsMobileNo(strInput) Then
            '103000：李南春，2017/2/7，按手机号查找
            If GetPatiIDFromCardType(mcnOracle, "手机号", strInput, False, lng病人ID, strCardPassWord, strErrMsg) = False Then lng病人ID = 0
            If lng病人ID = 0 Then GoTo NotFindPati:
            If GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg) = False Then Exit Function
        Else
            '按姓名全字匹配查找
            If GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "姓名", strInput) = False Then Exit Function
        End If
        GetPatient = True: Exit Function
    Case "医保号", "门诊号", "住院号", "就诊卡", "IC卡号", "手机号"
        If GetPatiIDFromCardType(mcnOracle, objCard.名称, strInput, False, lng病人ID, strCardPassWord, strErrMsg, lngCardTypeID) = False Then lng病人ID = 0
        If lng病人ID = 0 Then GoTo NotFindPati:
        If lngCardTypeID <> objCard.接口序号 And lngCardTypeID <> 0 Then
             Set objCard = GetIDKindCard(lngCardTypeID, CardTypeID)
        End If
        If GetPatiInforFromPatiID(mcnOracle, lng病人ID, objPati, strErrMsg) = False Then Exit Function
        objPati.卡号 = strInput: objPati.密码 = strCardPassWord
        GetPatient = True: Exit Function
    Case Else
            GoTo NotFindPati:
    End Select
    GoTo NotFindPati:
     Exit Function
errHandle:
    strErrMsg = Err.Description
NotFindPati:
    objPati.卡号 = strInput
End Function

Public Property Get GetPubOneCardObject() As clsPublicOneCard
    Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    Set GetPubOneCardObject = mobjPubOneCard
End Property
