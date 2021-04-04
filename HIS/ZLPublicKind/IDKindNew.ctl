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

Public Enum Mach_Mode   'ƥ�䷽ʽ
    CardTypeName = 0
    CardTypeID = 1
    CardTypeIndex = 2
End Enum
Public Enum IDKind_RegType
    Save_ע����Ϣ = 0
    Save_����ȫ�� = 1
    Save_����ģ�� = 2
    Save_˽��ȫ�� = 3
    Save_˽��ģ�� = 4
End Enum
Private mRegType As gRegType

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private mstrPopuCaption As String
Private mstrCardType As String  '��ǰ�����
Private mblnNotItemClick As Boolean
Private mcnOracle As ADODB.Connection
Private vRect As RECT
Private mobjCurCard As Card
Private mobjCards As Cards  '��Ч�����
Private mobjDefaultCard As Card '��ǰȱʡ�����
Private mlngDefaultCardID As Long 'ȱʡ�Ķ������ID
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

'ȱʡ����ֵ:
Const m_def_Locked = False
Const m_def_ProductName = "IDKindNew"
Const m_def_SaveRegType = IDKind_RegType.Save_˽��ȫ��
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
Const m_def_DefaultCardType = "���￨"
Const m_def_ShowPropertySet = False
Const m_def_IDKind = 0
Const m_def_CaptionAlignment = 2
'Const m_def_AutoSize = False
Const m_def_BorderStyle = IDKind_BorderStyle.ShowNone
Const m_def_IDKindStr = "��|��������￨|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;��|�ֻ���|0"
Const m_def_Appearance = 0
Const m_def_ShowSortName = True
'���Ա���:
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
'�¼�����:
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event Click(objCard As Object)
Public Event ItemClick(Index As Integer, objCard As Object)
Public Event ReadCard(ByVal objCard As Object, objPatiInfor As Object, blnCancel As Boolean)
Private mobjPubOneCard As clsPublicOneCard

Private Sub InitCardsObject()
    Dim strValue As String
    Dim objCard As Card, strCardTypes As String
    Dim i As Long, bln������ʾ As Boolean
    Dim DefaultCardType As Long
    
    On Error GoTo errHandle
    
    Set mobjCards = New Cards
    If Cards Is Nothing Then Exit Sub
    
     
    If mobjPubOneCard Is Nothing Then Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
 
    '78768:���ϴ�,2014/11/26,��������浽������
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("��ǰ����-���ܼ�", glngSys, 1153)
    End If
    
    If strValue = "" Then Call GetRegInFor(mRegType, "ҽ�ƿ����", "��ǰ����-���ܼ�", strValue)
    Cards.��ǰ�������ܼ� = strValue
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("��ǰ����-���", glngSys, 1153)
    End If
    If strValue = "" Then Call GetRegInFor(mRegType, "ҽ�ƿ����", "��ǰ����-���", strValue)
    If strValue = "" Then Cards.��ǰ�������ܼ� = "SHIFT"
    Cards.��ǰ������� = IIf(strValue = "", "F4", strValue)
    mobjCards.��ǰ�������ܼ� = Cards.��ǰ�������ܼ�
    mobjCards.��ǰ������� = Cards.��ǰ�������
    
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("������-���ܼ�", glngSys, 1153)
    End If
    If strValue = "" Then Call GetRegInFor(mRegType, "ҽ�ƿ����", "������-���ܼ�", strValue)
    Cards.���������ܼ� = strValue
    mobjCards.���������ܼ� = Cards.���������ܼ�
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("������-���", glngSys, 1153)
    End If
    If strValue = "" Then Call GetRegInFor(mRegType, "ҽ�ƿ����", "������-���", strValue)
    Cards.��������� = IIf(strValue = "", "F4", strValue)
    mobjCards.��������� = Cards.���������
    
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("����-���ܼ�", glngSys, 1153)
    End If
    If strValue = "" Then Call GetRegInFor(mRegType, "ҽ�ƿ����", "����-���ܼ�", strValue)
    Cards.�������ܼ� = strValue
    mobjCards.�������ܼ� = Cards.�������ܼ�
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("����-���", glngSys, 1153)
    End If
    If strValue = "" Then Call GetRegInFor(mRegType, "ҽ�ƿ����", "����-���", strValue)
    Cards.������� = IIf(strValue = "", "�ո��", strValue)
    mobjCards.������� = Cards.�������
    
    '78768:���ϴ�,2014/11/26,ȱʡ�������
    If Not mcnOracle Is Nothing Then
        strValue = mobjPubOneCard.getPara("ȱʡ�������", glngSys, 1153, 0)
    Else
        Call GetRegInFor(mRegType, "ҽ�ƿ����", "ȱʡ�������", strValue)
    End If
    mlngDefaultCardID = Val(strValue)
    
    Set mobjDefaultCard = Nothing
    With mobjCards
        For i = 1 To Cards.Count
            If m_ShowPropertySet And Cards(i).���� Then
                Call GetRegInFor(mRegType, "ҽ�ƿ����\" & Cards(i).����, "����", strValue)
                If strValue = "" Then   'ȱʡ����
                    Cards(i).���� = True
                Else
                    Cards(i).���� = Val(strValue) <> 0
                End If
            ElseIf Cards(i).�ӿ���� = 0 Then
                Call GetRegInFor(mRegType, "ҽ�ƿ����\" & Cards(i).����, "����", strValue)
                If strValue = "" Then   'ȱʡ����
                    Cards(i).���� = True
                Else
                    Cards(i).���� = Val(strValue) <> 0
                End If
            End If
            
            Call GetRegInFor(mRegType, "ҽ�ƿ����\" & Cards(i).����, "����-���ܼ�", strValue)
            Cards(i).���ܼ� = strValue
            Call GetRegInFor(mRegType, "ҽ�ƿ����\" & Cards(i).����, "����-���", strValue)
            Cards(i).��� = strValue
            'ȱʡ��ҽ�ƿ����
            '118959:���ϴ���2018/1/3,ȱʡ�������
            '1.��ҽ�ƿ����.ȱʡ��־=1��ҽ�ƿ�Ϊȱʡ���
            '2.��IDKindNew��DefaultCardTypeΪȱʡ���
            '3.�Ե�һ�����õ�ҽ�ƿ����Ϊȱʡҽ�ƿ����
            '76843:���ϴ�,2014/8/22,����ȱʡ��ҽ�ƿ�����
            If Cards(i).ȱʡ��־ Then
                DefaultCardType = Cards(i).�ӿ����
                Set mobjDefaultCard = Cards(i)
            End If
            If m_DefaultCardType <> "" And mobjDefaultCard Is Nothing And (Cards(i).���� = m_DefaultCardType _
                Or Cards(i).�ӿ���� = Val(m_DefaultCardType) And Val(m_DefaultCardType) <> 0) Then
                DefaultCardType = Cards(i).�ӿ����
                Set mobjDefaultCard = Cards(i)
            End If
            If Cards(i).�ӿ���� > 0 And objCard Is Nothing Then
                If objCard Is Nothing Then
                    DefaultCardType = Cards(i).�ӿ����
                    Set objCard = Cards(i)
                End If
            End If
            
            If Cards(i).���� Like "����*" Then
                Cards(i).ģ�������� = True
            End If
            If Cards(i).���� Then
                If Not bln������ʾ Then bln������ʾ = IIf(Cards(i).�������Ĺ��� <> "" And Cards(i).�������Ĺ��� <> "0", True, False)
                If Cards(i).�Ƿ�ģ������ And Cards(i).�ӿ���� > 0 Then
                    strCardTypes = strCardTypes & "," & Cards(i).�ӿ����
                End If
                If Cards(i).�ӿ���� = 0 Then
                    mobjCards.Add Cards(i), "M" & Cards(i).����
                Else
                    mobjCards.Add Cards(i), "K" & Cards(i).�ӿ����
                End If
            End If
        Next
        .��ȱʡ������ = Cards.��ȱʡ������
        If strCardTypes <> "" Then strCardTypes = Mid(strCardTypes, 2)
        .ģ��������� = strCardTypes
        .������ʾ = bln������ʾ
        If Not gobjCards Is Nothing Then
            gobjCards.������ʾ = bln������ʾ
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
    
    mstrPopuCaption = IIf(ShowSortName, "��", "����")
    m_IDKind = -1
    If mobjCards Is Nothing Then Exit Sub
    If mobjCards.Count = 0 Then Exit Sub
    
    Err = 0: On Error Resume Next
    If mstrCardType = "" Then
        If mobjCards.Count <> 0 Then
            mstrCardType = IIf(mobjCards(1).�ӿ���� > 0, mobjCards(1).�ӿ����, mobjCards(1).����)
        End If
    End If
    For i = 1 To mobjCards.Count
        strCardType = IIf(mobjCards(i).�ӿ���� > 0, mobjCards(i).�ӿ����, mobjCards(i).����)
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
    
    mstrPopuCaption = IIf(ShowSortName, mobjCurCard.����, mobjCurCard.����)
    If mstrCardType <> IIf(mobjCurCard.�ӿ���� > 0, mobjCurCard.�ӿ����, mobjCurCard.����) Then
        strCardType = IIf(mobjCurCard.�ӿ���� > 0, mobjCurCard.�ӿ����, mobjCurCard.����)
    End If
    picKind.ToolTipText = mobjCurCard.����
    If Ambient.UserMode And mblnNotItemClick = False Then
        RaiseEvent ItemClick(m_IDKind, mobjCurCard)
    End If
 End Sub
Private Sub SetCaption()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ñ���
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
        '78768:���ϴ�,2014/11/26,ȱʡ�������
        '85565,���ϴ�,2015/7/19:��������,�Ӵ�ʽ������Ҫ���
        ElseIf CheckReadCard Then
            intType = 2
        ElseIf mobjCurCard.�Ƿ�Ӵ�ʽ���� Then
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
    '����Kind�Ŀ��
    If Not AutoSize Then
        '���Զ�����
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
    '����:������
    '����:���˺�
    '����:2012-08-15 15:51:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True '����VisualTheme����Ч
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
    'intType:0-����;1-OnlyDown;2-OnlyKind
    If intType = 0 Or intType = 2 Then
        zlRaisEffect picKind, intStyle, mstrPopuCaption, intAlignMent
    End If
    If intType = 0 Or intType = 1 Then
        zlRaisEffect picDown, intStyle, " ", intAlignMent
    End If
End Sub
Private Sub CreatePopuMenu()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ʱ�˵�
    '����:���˺�
    '����:2012-11-21 09:49:35
    '˵��:
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
            Set mobjControl = .Add(xtpControlButton, mMenu_Kinds + j, objCard.����)
            strCardType = IIf(objCard.�ӿ���� > 0, objCard.�ӿ����, objCard.����)
            
            mobjControl.Parameter = strCardType
            If mstrCardType = strCardType Then
                mstrPopuCaption = IIf(ShowSortName, objCard.����, objCard.����)
                mstrCardType = strCardType
            End If
            j = j + 1
        Next
        
        If ShowPropertySet Then
            '��ʾ��������
            Set mobjControl = .Add(xtpControlButton, mMenu_Kinds + j, "�����������")
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
        '��������
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
    objPati.���� = strCardNo
    
    RaiseEvent ReadCard(objCard, objPati, blnCancel)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    Dim objPati As clsPatiInfor
    Dim blnCancel As Boolean
    Dim objCard As Card
    Set objCard = GetIDKindCard("IC��", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    Set objPati = New clsPatiInfor
    objPati.���� = strCardNo
    RaiseEvent ReadCard(objCard, objPati, blnCancel)
End Sub
Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    Dim objPati As clsPatiInfor
    Dim blnCancel As Boolean
    Dim objCard As Card
    Dim objStdPic As StdPicture
    Set objCard = GetIDKindCard("���֤", CardTypeName)
    If objCard Is Nothing Then Exit Sub
    Set objPati = New clsPatiInfor
    objPati.���� = strID
    objPati.���� = strName
    objPati.�Ա� = strSex
    objPati.�������� = Format(datBirthday, "yyyy-mm-DD HH:MM:SS")
    objPati.������ַ = strAddress
    objPati.���֤�� = strID
    objPati.���� = strNation
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    objPati.��Ƭ = objStdPic
    objPati.��Ƭ�ļ� = ""
    RaiseEvent ReadCard(objCard, objPati, blnCancel)
End Sub

Private Function IDKindClick(ByVal objCard As Card, Optional ByVal blnFastKeyUse As Boolean = False) As Boolean
    Dim objPati As New clsPatiInfor
    Dim blnCancel As Boolean, strErrMsg As String
    
    Dim strExpand As String, strOutCardNO As String, strOutPatiInforXML As String, strPhotoFile As String
    '���IDkindClick
    If objCard.���� Like "IC��*" And objCard.ϵͳ Then
        '72861:������,2014-05-09,��ݼ��޷�����IC���Ķ�������
        If mobjICCard Is Nothing Then
            'If blnFastKeyUse = True Then   '
                RaiseEvent Click(mobjCurCard)
                Call SetCaption
                Call ClearTag
                IDKindClick = True
            'End If
            Exit Function
        End If
        objPati.���� = mobjICCard.Read_Card()
        RaiseEvent ReadCard(objCard, objPati, blnCancel)
        If blnCancel = True Then Exit Function
        IDKindClick = True
        Exit Function
    End If
    Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    
    If objCard.�ӿ���� <= 0 Or mobjPubOneCard Is Nothing Then Exit Function
    
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOnlyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOnlyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��
    If mobjPubOneCard.objThirdSwap.zlReadCard(mobjParent, glngModul, objCard.�ӿ����, m_OnlyReadCardNo, strExpand, strOutCardNO, strOutPatiInforXML, strPhotoFile) = False Then Exit Function
    If Not m_OnlyReadCardNo Then
        Call zlGetPatiInforFromXML(mcnOracle, strOutPatiInforXML, objPati, strErrMsg)
        If objPati.��Ƭ Is Nothing And Trim(strPhotoFile) <> "" Then
            On Error Resume Next
            objPati.��Ƭ = LoadPicture(strPhotoFile)
            Err = 0: On Error GoTo 0
        End If
    End If
    If objPati Is Nothing Then Set objPati = New clsPatiInfor
    If objPati.���� = "" Then objPati.���� = strOutCardNO
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
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    '����:51488
    If Not mobjCurCard Is Nothing Then
         If mobjCurCard.�Ƿ�ˢ�� Or mobjCurCard.�Ƿ�ɨ�� Then Exit Sub
    End If
    If (m_Cards.������� = "�ո��" Or m_Cards.������� = " ") And Chr(KeyAscii) = " " Then KeyAscii = 0: Exit Sub
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
    '85565:���ϴ�,2015/7/23,��������
    '78768:���ϴ�,2014/11/26,ȱʡ�������
    If mobjCurCard.�Ƿ�Ӵ�ʽ���� = False And CheckReadCard = False Then Exit Sub
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
            '78768:���ϴ�,2014/11/26,ȱʡ�������
            ElseIf CheckReadCard Then
                SetCommandStatu 2, 0
            ElseIf mobjCurCard.�Ƿ�Ӵ�ʽ���� Then
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
    Dim objCurCard As Card, blnȱʡ���� As Boolean
    Dim lngPreIDKind As Long, intIndex As Integer
    If Button <> 1 Then Exit Sub
    If mobjCurCard Is Nothing Then Exit Sub
    '78768:���ϴ�,2014/11/26,ȱʡ�������
    blnȱʡ���� = CheckReadCard
    If mobjCurCard.�Ƿ�Ӵ�ʽ���� = False And blnȱʡ���� = False Then Exit Sub
    If blnȱʡ���� Then
        Set objCurCard = mobjCurCard
        Set mobjCurCard = Cards("K" & mlngDefaultCardID)
    End If
    If IDKindClick(mobjCurCard) = True Then
        If blnȱʡ���� Then Set mobjCurCard = objCurCard
        Call SetCaption
        Call ClearTag
        Exit Sub
    End If
'    RaiseEvent Click(mobjCurCard)
    If blnȱʡ���� Then Set mobjCurCard = objCurCard
    Call SetCaption
    Call ClearTag
End Sub
Private Sub ClearTag()
     shapBack.BorderColor = &H8000000A
     picKind.Tag = "": picDown.Tag = ""
End Sub

Private Sub UserControl_ExitFocus()
    'ƽ��,�ָ�ƽ��
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
    '����
      
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

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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
    '����:�жϴ�����Ƿ�Ϊ�ֻ���
    '���:strInput-�ֻ���
    '����:strRutType-��ѯ���:0-�ɹ�;1-������Ч�Ŷ�;2-���볤�Ȳ���
    '����:True-�������Ϊ�ֻ���;False-������벻Ϊ�ֻ���
    '����:������
    '����:2017-1-25
    '---------------------------------------------------------------------------------------------
    strRutType = 0
    'If mcnOracle Is Nothing Then IsMobileNo = False: Exit Function
    Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    IsMobileNo = mobjPubOneCard.zlIsMobileNo(strInput, strRutType)
    Exit Function
errHand:
    strRutType = 1
End Function


'Ϊ�û��ؼ���ʼ������
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
    Case Save_����ģ��
        mRegType = g����ģ��
    Case Save_����ȫ��
        mRegType = g����ȫ��
    Case Save_˽��ģ��
        mRegType = g˽��ģ��
    Case Save_˽��ȫ��
        mRegType = g˽��ȫ��
    Case Save_ע����Ϣ
        mRegType = gע����Ϣ
    Case Else
        mRegType = g˽��ȫ��
    End Select
End Sub
'�Ӵ������м�������ֵ
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
    '����:���õ�ǰ�ؼ��ĸ��������
    '����:���˺�
    '����:2018-12-19 10:02:20
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



'������ֵд���洢��
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
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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
        strNewIdKindStr = strNewIdKindStr & ";" & IIf(Cards(i).���� = "", Left(Cards(i).����, 1), Cards(i).����)
        strNewIdKindStr = strNewIdKindStr & "|" & Cards(i).����
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).�Ƿ�ˢ��, 0, 1)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).�ӿ���� < 0, 0, Cards(i).�ӿ����)
        strNewIdKindStr = strNewIdKindStr & "|" & Cards(i).���ų���
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).ȱʡ��־, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).�Ƿ�����ʻ�, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & Cards(i).�������Ĺ���
        '�����Ƿ�ɨ�裬�Ƿ�Ӵ�ʽ�������Ƿ�ǽӴ�ʽ��������
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).�Ƿ�ɨ��, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).�Ƿ�Ӵ�ʽ����, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).�Ƿ�ǽӴ�ʽ����, 1, 0)
    Next
    If strNewIdKindStr <> "" Then strNewIdKindStr = Mid(strNewIdKindStr, 2)
    m_IDKindStr = strNewIdKindStr
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=14
Public Sub Locale(Optional ByVal intSkip As Integer = 1)
    If intSkip > 0 Then
        '���
        If m_IDKind + intSkip > mobjCards.Count Then
            m_IDKind = m_IDKind + intSkip - mobjCards.Count
        Else
            m_IDKind = m_IDKind + intSkip
        End If
        If m_IDKind < 1 Or m_IDKind > mobjCards.Count Then
            m_IDKind = mobjCards.Count
        End If
    Else: Print
        '��ǰ
        If m_IDKind + intSkip < 1 Then
            m_IDKind = mobjCards.Count
        Else
            m_IDKind = m_IDKind + intSkip
        End If
    End If
    Err = 0: On Error Resume Next
    mstrCardType = IIf(mobjCards(m_IDKind).�ӿ���� > 0, mobjCards(m_IDKind).�ӿ����, mobjCards(m_IDKind).����)
    Call ReInitCards
    RaiseEvent ItemClick(m_IDKind, mobjCurCard)
'    DoEvents
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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


'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,1
Public Property Get CaptionAlignment() As IDKind_CaptionAlignment
    CaptionAlignment = m_CaptionAlignment
End Property
Public Property Let CaptionAlignment(ByVal New_CaptionAlignment As IDKind_CaptionAlignment)
    m_CaptionAlignment = New_CaptionAlignment
    PropertyChanged "CaptionAlignment"
    Call SetCaption
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=picDown,picDown,-1,Font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set picKind.Font = New_Font
    Set UserControl.Font = New_Font
    Set picDown.Font = New_Font
    PropertyChanged "Font"
    mblnNotItemClick = True '���ⴥ��ItemClick�¼�
    Call ReInitCards
    Call UserControl_Resize
    mblnNotItemClick = False
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=picKind,picKind,-1,FontBold
Public Property Get FontBold() As Boolean
    FontBold = picKind.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    picKind.FontBold() = New_FontBold
    UserControl.FontBold() = New_FontBold
    PropertyChanged "FontBold"
    mblnNotItemClick = True '���ⴥ��ItemClick�¼�
    Call ReInitCards
    Call UserControl_Resize
    mblnNotItemClick = False
    
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=picKind,picKind,-1,FontSize
Public Property Get FontSize() As Single
    FontSize = picKind.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    picKind.FontSize() = New_FontSize
    UserControl.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=picKind,picKind,-1,FontName
Public Property Get FontName() As String
    FontName = picKind.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    picKind.FontName() = New_FontName
    UserControl.FontName() = New_FontName
    PropertyChanged "FontName"
    mblnNotItemClick = True '���ⴥ��ItemClick�¼�
    Call ReInitCards
    Call UserControl_Resize
    mblnNotItemClick = False
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=picKind,picKind,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = picKind.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    picKind.ForeColor() = New_ForeColor
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    mblnNotItemClick = True '���ⴥ��ItemClick�¼�
    Call ReInitCards
    Call UserControl_Resize
    mblnNotItemClick = False

End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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
    mstrCardType = IIf(mobjCurCard.�ӿ���� > 0, mobjCurCard.�ӿ����, mobjCurCard.����)
    mstrPopuCaption = IIf(ShowSortName, mobjCurCard.����, mobjCurCard.����)
    picKind.ToolTipText = mobjCurCard.����
    Call SetCaption
    RaiseEvent ItemClick(m_IDKind, objCard)
'    DoEvents
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=14
Public Function GetIDKindCard(ByVal strCardType As String, _
    Optional MachMode As Mach_Mode, Optional bln���ѿ� As Boolean = False) As Object
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����
    '���:strCardType-�����(����Ϊָ���Ŀ����ID,�ַ�Ϊ������ƥ��)
    '     MachMode-ƥ�䷽ʽ
    '     bln���ѿ�-�Ƿ����ѿ�����
    '����: �ɹ������ؿ�������;���򷵻�Nothing
    '����:���˺�
    '����:2012-08-20 18:20:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, i As Long
    If MachMode <> CardTypeName Then
        lngCardTypeID = Val(strCardType)
        Err = 0: On Error Resume Next
        If MachMode = CardTypeID Then
            For i = 1 To Cards.Count
                If Cards(i).�ӿ���� = lngCardTypeID And Cards(i).���ѿ� = bln���ѿ� Then
                    Set GetIDKindCard = Cards(i): Exit Function
                End If
            Next
            Set GetIDKindCard = Nothing
            Exit Function
        Else
            If lngCardTypeID = -1 Then Set GetIDKindCard = Nothing: Exit Function
            Set GetIDKindCard = mobjCards(lngCardTypeID)     '����ֻ��ȡ��Ч�Ŀ�������
        End If
        If Err <> 0 Then Set GetIDKindCard = Nothing
        Exit Function
    End If
    
    For i = 1 To Cards.Count
        Select Case strCardType
        Case "���֤", "���֤��", "�������֤"
            If InStr(1, Cards(i).����, "���֤") > 0 Then
                 Set GetIDKindCard = Cards(i): Exit Function
            End If
        Case Else
            If strCardType = Cards(i).���� Then
                 Set GetIDKindCard = Cards(i): Exit Function
            End If
        End Select
    Next
    Set GetIDKindCard = Nothing
End Function


'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=14
Public Function GetKindIndex(ByVal strCardType As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ����
    '���:strCardType-�����(����Ϊָ���Ŀ����ID,�ַ�Ϊ������ƥ��)
    '����: �ɹ�����������ֵ,���򷵻�-1
    '����:���˺�
    '����:2012-08-20 18:20:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCardTypeID As Long, i As Long
    Dim blnCardTypeID As Boolean
    
    blnCardTypeID = IsNumeric(strCardType)
    lngCardTypeID = Val(strCardType)
    For i = 1 To mobjCards.Count
        If blnCardTypeID Then
            If mobjCards(i).�ӿ���� = lngCardTypeID Then GetKindIndex = i: Exit Function
        Else
            Select Case strCardType
            Case "���֤", "���֤��", "�������֤"
                If InStr(1, mobjCards(i).����, "���֤") > 0 Then
                     GetKindIndex = i: Exit Function
                End If
            Case "IC��", "IC����"
                If InStr(1, mobjCards(i).����, "IC��") > 0 Then
                     GetKindIndex = i: Exit Function
                End If
                
            Case Else
                If strCardType Like "����*" And mobjCards(i).���� Like "����*" Then
                         GetKindIndex = i: Exit Function
                Else
                    If strCardType = mobjCards(i).���� Then
                         GetKindIndex = i: Exit Function
                    End If
                End If
            End Select
        End If
    Next
    GetKindIndex = -1
End Function
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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
    '����:��ȡȱʡ��
    '����:���˺�
    '����:2012-08-24 10:26:54
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


'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,False
Public Property Get SmallStyle() As Boolean
    SmallStyle = m_SmallStyle
End Property

Public Property Let SmallStyle(ByVal New_SmallStyle As Boolean)
    Dim MyFont As New StdFont
    'Ϊ�����ϵ�KindǸ��
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

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=8,0,0,2
Public Property Get KeyShift() As Long
    KeyShift = m_KeyShift
End Property

Public Property Let KeyShift(ByVal New_KeyShift As Long)
    If New_KeyShift < 256 Then
        m_KeyShift = New_KeyShift
        PropertyChanged "KeyShift"
    Else
        MsgBox "��Ч������ֵ,��ο�KeyDown�¼���shift����", vbInformation, App.ProductName
    End If
End Property

Private Sub mobjParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If UserControl.Enabled Then
        If Shift = KeyShift Then
            'keycode:96��С���̵�0,105��9
            If KeyCode > 95 And KeyCode < 106 Then
                IDKind = KeyCode - 96
            ElseIf KeyCode = 123 Then   'Ctrol+F12,������ִ�е�������Ķ�������Ӧ���ȼ�
                Call picKind_MouseDown(1, Shift, 0, 0)
            End If
        End If
    End If
End Sub
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7
Public Function ListCount() As Integer
    If mobjCards Is Nothing Then Exit Function
    ListCount = mobjCards.Count
End Function

Public Function MacthKey(strFunKey As String, strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ƥ�䰴ť�Ƿ���ȷ
    '���:KeyCode-����
    '       Shift-���ܼ�
    '����: ƥ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-23 09:32:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intFun As Integer
    Dim kbState(255) As Byte, lngReturn As Long
    Dim blnDownCTRL As Boolean
    Dim blnSHIFT As Boolean
    Dim blnFun As Boolean '���ܼ��Ƿ�����
    Dim blnTemp As Boolean
    
    lngReturn = GetKeyboardState(kbState(0))
    If lngReturn = 0 Then Exit Function
    
    blnDownCTRL = kbState(vbKeyControl) And &H80
    blnSHIFT = kbState(vbKeyShift) And &H80
    
    strFunKey = UCase(strFunKey)
    strKey = UCase(strKey)
    
    blnFun = IIf(strFunKey = "CTRL", blnDownCTRL, IIf(strFunKey = "SHIFT", blnSHIFT, Not (blnDownCTRL Or blnSHIFT)))
    
    If Left(strKey, 1) = "F" And Len(strKey) >= 2 Then
        'F1-Fn��ͷ
        i = vbKeyF1 - 1 + Val(Mid(strKey, 2))
        MacthKey = kbState(i) And &H80 And blnFun
 
        Exit Function
    End If
 
    If strKey = " " Or strKey = "�ո��" Then
        MacthKey = kbState(vbKeySpace) And &H80 And blnFun
        Exit Function
    End If
         
 
    '��������
    If InStr(" ��������", strKey) > 0 Then
        i = Switch(strKey = "��", vbKeyLeft, strKey = "��", vbKeyUp, strKey = "��", vbKeyRight, True, vbKeyDown)
        MacthKey = kbState(i) And &H80 And blnFun
        Exit Function
    End If
   
    i = Asc(strKey)
    If i >= Asc("A") And i <= Asc("Z") Then
        '��ĸ
        i = i - Asc("A")
        i = vbKeyA + i
        MacthKey = kbState(i) And &H80 And blnFun
        Exit Function
    End If
    
    'С���̵ļ�(���ּ���)
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
    Case UCase("KeyLButton")    'vbKeyLButton 0x1 ������
        blnTemp = kbState(vbKeyLButton) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyMButton")   'vbKeyRButton 0x2 ����Ҽ�'
        blnTemp = kbState(vbKeyRButton) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyMButton")   'vbKeyMButton 0x4 ����м�
        blnTemp = kbState(vbKeyMButton) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyCancel")   ' vbKeyCancel 0x3 CANCEL ��
        blnTemp = kbState(vbKeyCancel) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyBack")   ' vbKeyBack 0x8 BACKSPACE ��
        blnTemp = kbState(vbKeyBack) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyTab")   ' vbKeyTab 0x9 TAB ��
        blnTemp = kbState(vbKeyTab) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyClear")   ' vbKeyClear 0xC CLEAR ��
        blnTemp = kbState(vbKeyClear) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("ENTER")   ' vbKeyReturn 0xD ENTER ��
        blnTemp = kbState(vbKeyReturn) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("SHIFT")   ' vbKeyShift 0x10 SHIFT ��
        blnTemp = kbState(vbKeyShift) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("CTRL")   ' vbKeyControl 0x11 CTRL ��
        blnTemp = kbState(vbKeyControl) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyMenu")   ' vbKeyMenu 0x12 MENU ��
        blnTemp = kbState(vbKeyMenu) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("KeyPause")   ' vbKeyPause 0x13 PAUSE ��
        blnTemp = kbState(vbKeyPause) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase(" CAPS LOCK")   ' vbKeyCapital 0x14 CAPS LOCK ��
        blnTemp = kbState(vbKeyCapital) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("ESC")   ' vbKeyEscape 0x1B ESC ��
        blnTemp = kbState(vbKeyEscape) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("Space")   ' vbKeySpace 0x20 SPACEBAR ��
        blnTemp = kbState(vbKeySpace) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("SELECT")   ' vbKeySelect 0x29 SELECT ��
        blnTemp = kbState(vbKeySelect) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("PRINT SCREEN")   ' vbKeyPrint 0x2A PRINT SCREEN ��
        blnTemp = kbState(vbKeyPrint) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("EXECUTE")   ' vbKeyExecute 0x2B EXECUTE ��
        blnTemp = kbState(vbKeyPrint) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("SNAPSHOT")   ' vbKeySnapshot 0x2C SNAPSHOT ��
        blnTemp = kbState(vbKeySnapshot) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("INSERT")   ' vbKeyInsert 0x2D INSERT ��
        blnTemp = kbState(vbKeyInsert) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("DELETE")   ' vbKeyDelete 0x2E DELETE ��
        blnTemp = kbState(vbKeyDelete) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("HELP")   ' vbKeyHelp 0x2F HELP ��
        blnTemp = kbState(vbKeyHelp) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case UCase("NUM LOCK")   ' vbKeyNumlock 0x90 NUM LOCK ��
        blnTemp = kbState(vbKeyNumlock) And &H80: MacthKey = blnTemp And blnFun: Exit Function
    Case Else
    End Select
End Function

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "����/����һ��ֵ������һ�������Ƿ���Ӧ�û������¼���"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    picDown.Enabled = New_Enabled
    picDown.Enabled = New_Enabled
    
    PropertyChanged "Enabled"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,False
Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    mblnNotItemClick = True '���ⴥ��ItemClick�¼�
    Call ReInitCards
    Call UserControl_Resize
    mblnNotItemClick = False

End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=9,0,0,0
Public Property Get objParent() As Object
    Set objParent = mobjParent
End Property
Public Property Set objParent(ByVal New_objParent As Object)
    Set mobjParent = New_objParent
    PropertyChanged "objParent"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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
    '����:��ʼ��
    '���:frmMain-���õ�������
    '     lngSys : ϵͳ���
    '     lngModul:��Ҫִ�еĹ������
    '     objCardSquare-�����㲿��
    '     cnOracle:����������ݿ�����
    '     strIDKindStr-���ʶ��������,�����ָ�ʽ:
    '               һ����ȱʡ��:����1|ȫ��1|������־1;��. ;����n|ȫ��n|������־n
    '               һ������չ��ʽ:����1|ȫ��1|������־1|�����ID1|���ų���1|ȱʡ��־1(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�1(1-�����ʻ�;0-�������ʻ�)|��������1(�ڼ�λ���ڼ�λ����,��Ϊ������);��
    '   strProductName-��Ʒģ������(��Ҫ�������ģ������ı���)
    '    blnOnlyThreeCard-��������
    '    blnIsObjRegisterAlone-�Ƿ�ʹ�ö�����ע�Ჿ��(True:ʹ��:zlRegisterAlone.DLL,����ʹ��zlRegister.dll)
    '����:���˺�
    '����:2012-08-16 10:45:14
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
    '������������
    Call CreateComEvtsObject
    '80115:���ϴ�,2014/12/1,��ʼ����ǰ�����
    mstrCardType = "": Call ReInitCards
    
End Sub

Private Sub CreateComEvtsObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������¼�����
    '����: �����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-08-28 16:16:00
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    If mobjPubOneCard Is Nothing Then Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    '������������
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
    '�ر���ض���
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

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=14
Public Sub SetAutoReadCard(ByVal blnAutoReadCard As Boolean)
    '�����Զ�����
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

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,False
Public Property Get AllowAutoICCard() As Boolean
    AllowAutoICCard = m_AllowAutoICCard
End Property

Public Property Let AllowAutoICCard(ByVal New_AllowAutoICCard As Boolean)
    m_AllowAutoICCard = New_AllowAutoICCard
    PropertyChanged "AllowAutoICCard"
    '76256,Ƚ����,2014-8-5,�ȳ�ʼ��IDKind(zlInit),�����ø�����,δ���������¼�����
    '������������
    Call CreateComEvtsObject
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,False
Public Property Get AllowAutoIDCard() As Boolean
    AllowAutoIDCard = m_AllowAutoIDCard
End Property

Public Property Let AllowAutoIDCard(ByVal New_AllowAutoIDCard As Boolean)
    m_AllowAutoIDCard = New_AllowAutoIDCard
    PropertyChanged "AllowAutoIDCard"
    '76256,Ƚ����,2014-8-5,�ȳ�ʼ��IDKind(zlInit),�����ø�����,δ���������¼�����
    '������������
    Call CreateComEvtsObject
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,True
Public Property Get AllowAutoCommCard() As Boolean
    AllowAutoCommCard = m_AllowAutoCommCard
End Property

Public Property Let AllowAutoCommCard(ByVal New_AllowAutoCommCard As Boolean)
    m_AllowAutoCommCard = New_AllowAutoCommCard
    PropertyChanged "AllowAutoCommCard"
    '76256,Ƚ����,2014-8-5,�ȳ�ʼ��IDKind(zlInit),�����ø�����,δ���������¼�����
    '������������
    Call CreateComEvtsObject
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,False
Public Property Get NotAutoAppendKind() As Boolean
    NotAutoAppendKind = m_NotAutoAppendKind
End Property

Public Property Let NotAutoAppendKind(ByVal New_NotAutoAppendKind As Boolean)
    m_NotAutoAppendKind = New_NotAutoAppendKind
    PropertyChanged "NotAutoAppendKind"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,1,0,False
Public Property Get DefaultCardShowPassText() As Boolean
    If m_Cards.��ȱʡ������ And Not mobjDefaultCard Is Nothing Then
        DefaultCardShowPassText = mobjDefaultCard.�������Ĺ��� <> ""
    Else
        If mobjCards Is Nothing Then
             DefaultCardShowPassText = Cards.������ʾ
        Else
             DefaultCardShowPassText = mobjCards.������ʾ
        End If
    End If
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,1,0,False
Public Property Get ShowPassText() As Boolean
    If mobjCurCard Is Nothing Then Exit Property
    If mobjCurCard.ģ�������� Then
        ShowPassText = DefaultCardShowPassText
    Else
         ShowPassText = IIf(mobjCurCard.�������Ĺ��� = "0", "", mobjCurCard.�������Ĺ���) <> ""
    End If
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,1,0,100
Public Property Get GetDefaultCardNoLen() As Integer
    If m_Cards Is Nothing Then GetDefaultCardNoLen = 100: Exit Property
    If m_Cards.��ȱʡ������ And Not mobjDefaultCard Is Nothing Then
        GetDefaultCardNoLen = mobjDefaultCard.���ų���
    Else
        GetDefaultCardNoLen = 100
    End If
End Property
Public Property Get GetCardNoLen() As Integer
    If mobjCurCard Is Nothing Then Exit Property
    If mobjCurCard.ģ�������� Then
        GetCardNoLen = GetDefaultCardNoLen
    Else
         GetCardNoLen = mobjCurCard.���ų���
    End If
End Property

Private Function ActiveFastKeyInside() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������
    '����: ��������,����True,���򷵻�False
    '����:���˺�
    '����:2012-08-29 22:58:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '������
    Dim objCurCard As Card, blnȱʡ���� As Boolean
    blnȱʡ���� = CheckReadCard
    If mobjCurCard Is Nothing Or m_Cards Is Nothing Then Exit Function
    '�������IDKind
     If m_Cards.��ǰ�������ܼ� <> "" Or m_Cards.��ǰ������� <> "" Then
            If MacthKey(m_Cards.��ǰ�������ܼ�, m_Cards.��ǰ�������) Then
                Call Locale(-1): ActiveFastKeyInside = True: Exit Function
            End If
    Else
        'ȱʡ
        If MacthKey("SHIFT", "F4") Then Locale (-1): ActiveFastKeyInside = True: Exit Function
        'ȱʡ
        If MacthKey("", "��") Then Locale (-1): ActiveFastKeyInside = True: Exit Function
        If MacthKey("", "PgUp") Then Locale (-1): ActiveFastKeyInside = True: Exit Function
     End If
     If m_Cards.���������ܼ� <> "" Or m_Cards.��������� <> "" Then
            If MacthKey(m_Cards.���������ܼ�, m_Cards.���������) Then
                Call Locale(1): ActiveFastKeyInside = True: Exit Function
            End If
    Else
        'ȱʡ
        If MacthKey("", "F4") Then Locale (1): ActiveFastKeyInside = True: Exit Function
        'ȱʡ
        If MacthKey("", "��") Then Locale (1): ActiveFastKeyInside = True: Exit Function
        If MacthKey("", "PgDn") Then Locale (-1): ActiveFastKeyInside = True: Exit Function
     End If
    '����
    If mobjCurCard.�Ƿ�Ӵ�ʽ���� = False And blnȱʡ���� = False Then Exit Function
    If MacthKey(m_Cards.�������ܼ�, m_Cards.�������) = False Then Exit Function
    If blnȱʡ���� Then
        Set objCurCard = mobjCurCard
        Set mobjCurCard = Cards("K" & mlngDefaultCardID)
    End If
    Call IDKindClick(mobjCurCard, True)
    If blnȱʡ���� Then Set mobjCurCard = objCurCard
    ActiveFastKeyInside = True
End Function

Public Function ActiveFastKey() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������
    '����: ��������,����True,���򷵻�False
    '����:���˺�
    '����:2012-08-29 22:58:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mblnTextFocus = False Then
        ActiveFastKey = ActiveFastKeyInside
    End If
End Function
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=0,0,0,True
Public Property Get OnlyReadCardNo() As Boolean
    OnlyReadCardNo = m_OnlyReadCardNo
End Property
Public Property Let OnlyReadCardNo(ByVal New_OnlyReadCardNo As Boolean)
    m_OnlyReadCardNo = New_OnlyReadCardNo
    PropertyChanged "OnlyReadCardNo"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "����/���ö������ı���ͼ�εı���ɫ��"
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
    '����:��ȡȱʡ�Ŀ����ID
    '����: ȱʡ�Ŀ����ID;-1��ʾ��ģ�����ҽ��в��Ҳ���
    '����:���˺�
    '����:2012-08-31 17:02:08
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng�����ID As Long
    If m_Cards.��ȱʡ������ And Not mobjDefaultCard Is Nothing Then
        lng�����ID = mobjDefaultCard.�ӿ����
    Else
        lng�����ID = "-1"
    End If
    '������ɾ����,���²���
    GetDefaultCardTypeID = lng�����ID
End Function
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,3
Public Property Get SaveRegType() As IDKind_RegType
    SaveRegType = m_SaveRegType
End Property

Public Property Let SaveRegType(ByVal New_SaveRegType As IDKind_RegType)
    m_SaveRegType = New_SaveRegType
    Call FromRegType(m_SaveRegType)
    PropertyChanged "SaveRegType"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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
    '����:��ȡ��Ч��IDKindStr
    '����:��Ч��IDkindStr:����|ȫ��|��������|�ӿ����|���ų���|ȱʡ��־|�Ƿ�����ʻ�|�������Ĺ���;...
    '����:���˺�
    '����:2012-10-22 11:55:31
    '˵��:(Ŀǰ��ҪӦ�����Զ�����)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim strNewIdKindStr As String
    Dim i As Long
    
    If mobjCards Is Nothing Then Exit Function
    strNewIdKindStr = ""
    For i = 1 To mobjCards.Count
        Set objCard = mobjCards(i)
        strNewIdKindStr = strNewIdKindStr & ";" & IIf(objCard.���� = "", Left(objCard.����, 1), objCard.����)
        strNewIdKindStr = strNewIdKindStr & "|" & objCard.����
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).�Ƿ�ˢ��, 0, 1)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(objCard.�ӿ���� < 0, 0, objCard.�ӿ����)
        strNewIdKindStr = strNewIdKindStr & "|" & objCard.���ų���
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(objCard.ȱʡ��־, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(objCard.�Ƿ�����ʻ�, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & objCard.�������Ĺ���
        '�����Ƿ�ɨ�裬�Ƿ�Ӵ�ʽ�������Ƿ�ǽӴ�ʽ��������
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).�Ƿ�ɨ��, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).�Ƿ�Ӵ�ʽ����, 1, 0)
        strNewIdKindStr = strNewIdKindStr & "|" & IIf(Cards(i).�Ƿ�ǽӴ�ʽ����, 1, 0)
    Next
    If strNewIdKindStr <> "" Then strNewIdKindStr = Mid(strNewIdKindStr, 2)
    AvailabilityIdkindStr = strNewIdKindStr
End Property
 
Public Property Get AvailabilityCards() As Object
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ч��Cards����
    '����:��Ч��Cards����
    '����:���˺�
    '����:2012-10-22 11:55:31
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set AvailabilityCards = mobjCards
End Property
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Private Function CheckReadCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ����ȱʡ�Ķ������
    '����: true-��������Ч��false-������
    '����:���ϴ�
    '����:2014/11/26
    '˵��:
    '����:78768
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCurCard As Card
    If (mobjCurCard.���� = "����" Or mobjCurCard.���� = "��������￨") And mlngDefaultCardID > 0 Then
        Err = 0: On Error Resume Next
        Set objCurCard = Cards("K" & mlngDefaultCardID)
        If Err <> 0 Then Err = 0: Exit Function
        If Not (objCurCard.�Ƿ�Ӵ�ʽ���� Or objCurCard.�Ƿ�ǽӴ�ʽ����) Then Exit Function
        CheckReadCard = True
    End If
End Function
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
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
    '����:����ˢ���ӿ�
    '����: true-�ɹ���false-ʧ��
    '����:���ϴ�
    '����:2015/7/23 10:37:39
    '����:85565
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    Dim objCard As Card
    
    Err = 0: On Error Resume Next
    SetBrushCardObject = True
    '118959:���ϴ�,2018/1/3,���֧����Ƶ���������¼�����
    If (mobjCurCard Is Nothing And mobjDefaultCard Is Nothing) Or mobjTxtInput Is Nothing Then Exit Function
    Set objCard = mobjCurCard
    If mobjCurCard.�ӿ���� = 0 And mobjCurCard.���� Like "*����*" Then Set objCard = mobjDefaultCard

    If objCard.�ӿ���� = 0 Or objCard.�ӿڳ����� = "" Or Not (objCard.�Ƿ�ˢ�� Or objCard.�Ƿ�ɨ��) Then Exit Function
    
    If mobjPubOneCard Is Nothing Then Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    If mobjPubOneCard Is Nothing Then Exit Function
    
    If mobjPubOneCard.objThirdSwap.zlSetBrushCardObject(objCard.�ӿ����, IIf(blnComm, mobjTxtInput, Nothing), strExpend, objCard.���ѿ�) Then
        If mobjCommEvents Is Nothing And AllowAutoCommCard = True Then
            Set mobjCommEvents = New clsCommEvents
            Call mobjPubOneCard.objThirdSwap.zlInitEvents(UserControl.hWnd, mobjCommEvents)
        End If
    End If
End Function

Public Function zlFindPatient(ByVal strInput As String, objPatiInfor As Object, _
    Optional objCard As Object, Optional strErrMsg As String, Optional ByVal blnBrushCard As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ͨ�������ֵ��ˢ���������ֵ���Ҳ��˲����ز�����Ϣ
    '���:strInput-�����ֵ��ˢ���������ֵ
    '     objCard-��ǰ�Ŀ����nothingʱ��ΪIDKindNew. CurCard�������Դ���Ϊ׼
    '     blnBrushCard-�Ƿ���ˢ��
    '����: objPati-��ǰ�����Ĳ�����Ϣ
    '      strErrMsg-���ش�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2012-09-24 15:13:43
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set objPatiInfor = New clsPatiInfor
    If strInput = "" And mobjTxtInput Is Nothing Then Exit Function
    If strInput = "" And Not mobjTxtInput Is Nothing Then strInput = Trim(mobjTxtInput.Text)
    If strInput = "" Then Exit Function
    zlFindPatient = GetPatient(strInput, objPatiInfor, objCard, strErrMsg)
End Function

Public Function zlGetPatiInforFromPatiID(ByVal lng����ID As Long, ByRef objPatiInfor As Object, Optional ByRef strErrMsg As String, Optional strOtherName As String = "", Optional strOtherValue As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ͨ������ID��ȡ������Ϣ
    '���:lng����ID
    '����: objPati-��ǰ�����Ĳ�����Ϣ
    '      strErrMsg-���ش�����Ϣ
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2012-09-24 15:13:43
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set objPatiInfor = New clsPatiInfor
    If lng����ID = 0 Then Exit Function
    zlGetPatiInforFromPatiID = GetPatiInforFromPatiID(mcnOracle, lng����ID, objPatiInfor, strErrMsg, strOtherName, strOtherValue)
End Function

Public Function zlGetPatiIDFromCardType(ByVal strCardType As String, ByVal strCardNo As String, _
    Optional ByVal blnNotShowErrMsg As Boolean = False, Optional ByRef lng����ID As Long, _
    Optional ByRef strCardPassWord As String, Optional ByRef strErrMsg As String, _
    Optional ByRef lngCardTypeID As Long, Optional objCtl As Object = Nothing, Optional frmMain As Object, _
    Optional blnShowMergePati As Boolean = False, Optional ByRef blnOnlyContractPati As Boolean = False, _
    Optional ByRef blnCertificate As Boolean = False, Optional ByRef blnUserCancel As Boolean = False, _
    Optional ByVal lngShowCardNoTypeID As Long = 0, Optional ByVal blnNotCheckValidDate As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ����ҽ�����Ϳ���,��ȡ��Ӧ�Ĳ���ID
    '���:strCardType-�����,���Ϊ����,��Ϊ�����ID,���Ϊ�ַ�,��Ϊ�������
    '       strCardNo-����
    '       blnNotShowErrMsg-����ʾ�������ʾ��Ϣ
    '       frmMain-���õ�������
    '       objCtl-���õĿؼ�
    '       blnShowMergePati-�����ֶ�����������Ĳ���ʱ,�Ƿ���ʾ�ϲ����ܰ�ť
    '       blnOnlyContractPati-ǩԼ����
    '       blnUserCancel-ѡ�����У��û�ѡ����ȡ��
    '       lngShowCardNoTypeID-���˳���������Ϣʱ������ѡ��������ʾ�Ŀ��ŵĿ����ID,0-��ʾ����ʾ���ţ�>0��ʾ��ʾָ����������ID
    '       blnNotCheckValidDate-�Ƿ�Կ���ֹʹ��ʱ����м��,true-�������ֹʹ��ʱ��,false-���
    '����:strErrMsg-���صĴ�����Ϣ
    '       lng����ID-���صĲ���ID
    '       strCardPass-���ؿ��ŵ�����
    '       lngCardTypeID-���ؿ����ID(0��ʾ����ȷ�������ID)
    '����:��ȡ����ID�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-14 17:07:51
    '˵��:ֻ�д���ҽ�����Ĳŵ��ô˺���
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetPatiIDFromCardType = GetPatiIDFromCardType(mcnOracle, strCardType, strCardNo, blnNotShowErrMsg, lng����ID, _
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
    Optional ByVal lngDefaultCardTypeID As Long, Optional ByVal blnȱʡ������ As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ID
    '���:strInput-��������ֵ
    '     objCard-��ǰ��ȡ�Ŀ����(����Ϊ����)
    '     blnBrushCard-�Ƿ���ˢ��
    '����:lng����ID-����:����ID
    '     objCard-���ض�ȡ�Ŀ����
    '     strErrMsg-���صĴ�����Ϣ
    '����: �ɹ�,����ָ���Ĳ���ID
    '����:���˺�
    '����:2012-08-20 17:34:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardPassWord As String, lngCardTypeID As Long
    Dim lng����ID As Long
    
    On Error GoTo errHandle
    
    'ָ�������
    If objCard Is Nothing Then Set objCard = GetCurCard
    If objCard.�ӿ���� > 0 Then
        If GetPatiIDFromCardType(mcnOracle, objCard.�ӿ����, strInput, False, lng����ID, strCardPassWord, strErrMsg, lngCardTypeID) Then
            If lngCardTypeID <> objCard.�ӿ���� And lngCardTypeID <> 0 Then
                 Set objCard = GetIDKindCard(lngCardTypeID, CardTypeID)
            End If
        ElseIf IsMobileNo(strInput) Then
            '103000�����ϴ���2017/2/7�����ֻ��Ų���
            If GetPatiIDFromCardType(mcnOracle, "�ֻ���", strInput, False, lng����ID, strCardPassWord, strErrMsg) Then
                Set objCard = GetIDKindCard("�ֻ���", CardTypeName)
            Else
                lng����ID = 0
            End If
        Else
            lng����ID = 0
        End If
        If lng����ID = 0 Then GoTo NotFindPati:
        If GetPatiIDFromCardType(mcnOracle, lng����ID, objPati, strErrMsg) = False Then Exit Function
        objPati.���� = strInput: objPati.���� = strCardPassWord
        GetPatient = True: Exit Function
    End If
    If objCard.���� Like "����*" Or objCard.ģ�������� Then
        If blnȱʡ������ Then
            lngCardTypeID = lngDefaultCardTypeID
        Else
            lngCardTypeID = -1
        End If
        If blnBrushCard Then
            If GetPatiIDFromCardType(mcnOracle, lngCardTypeID, strInput, False, lng����ID, strCardPassWord, strErrMsg, _
                lngCardTypeID) = False Then lng����ID = 0
            If lngCardTypeID <> objCard.�ӿ���� And lngCardTypeID <> 0 Then
                 Set objCard = GetIDKindCard(lngCardTypeID, CardTypeID)
            ElseIf Not GetfaultCard Is Nothing Then
                Set objCard = GetfaultCard
            End If
            If lng����ID = 0 Then GoTo NotFindPati:
            If GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg) = False Then Exit Function
            objPati.���� = strInput: objPati.���� = strCardPassWord
            GetPatient = True: Exit Function
        End If
    End If
    Select Case objCard.����
    Case "����", "��������￨"
        If Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '����ID
            lng����ID = Val(Mid(strInput, 2))
            If GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg) = False Then Exit Function
        ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
            If GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "�����", Mid(strInput, 2)) = False Then Exit Function
        ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
            If GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "סԺ��", Mid(strInput, 2)) = False Then Exit Function
        ElseIf IsMobileNo(strInput) Then
            '103000�����ϴ���2017/2/7�����ֻ��Ų���
            If GetPatiIDFromCardType(mcnOracle, "�ֻ���", strInput, False, lng����ID, strCardPassWord, strErrMsg) = False Then lng����ID = 0
            If lng����ID = 0 Then GoTo NotFindPati:
            If GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg) = False Then Exit Function
        Else
            '������ȫ��ƥ�����
            If GetPatiInforFromPatiID(mcnOracle, 0, objPati, strErrMsg, "����", strInput) = False Then Exit Function
        End If
        GetPatient = True: Exit Function
    Case "ҽ����", "�����", "סԺ��", "���￨", "IC����", "�ֻ���"
        If GetPatiIDFromCardType(mcnOracle, objCard.����, strInput, False, lng����ID, strCardPassWord, strErrMsg, lngCardTypeID) = False Then lng����ID = 0
        If lng����ID = 0 Then GoTo NotFindPati:
        If lngCardTypeID <> objCard.�ӿ���� And lngCardTypeID <> 0 Then
             Set objCard = GetIDKindCard(lngCardTypeID, CardTypeID)
        End If
        If GetPatiInforFromPatiID(mcnOracle, lng����ID, objPati, strErrMsg) = False Then Exit Function
        objPati.���� = strInput: objPati.���� = strCardPassWord
        GetPatient = True: Exit Function
    Case Else
            GoTo NotFindPati:
    End Select
    GoTo NotFindPati:
     Exit Function
errHandle:
    strErrMsg = Err.Description
NotFindPati:
    objPati.���� = strInput
End Function

Public Property Get GetPubOneCardObject() As clsPublicOneCard
    Call zlGetPubOneCard(mcnOracle, mobjPubOneCard)
    Set GetPubOneCardObject = mobjPubOneCard
End Property
