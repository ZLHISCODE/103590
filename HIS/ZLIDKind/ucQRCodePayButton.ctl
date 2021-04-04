VERSION 5.00
Begin VB.UserControl ucQRCodePayButton 
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1530
   ScaleHeight     =   570
   ScaleWidth      =   1530
   ToolboxBitmap   =   "ucQRCodePayButton.ctx":0000
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
   Begin VB.Shape shapBack 
      BorderColor     =   &H8000000A&
      Height          =   465
      Left            =   30
      Top             =   0
      Width           =   945
   End
   Begin VB.Image imgHot 
      Height          =   240
      Left            =   690
      Picture         =   "ucQRCodePayButton.ctx":0312
      Top             =   135
      Width           =   240
   End
   Begin VB.Image ImgNormal 
      Height          =   240
      Left            =   1155
      Picture         =   "ucQRCodePayButton.ctx":089C
      Top             =   165
      Width           =   240
   End
End
Attribute VB_Name = "ucQRCodePayButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************************************************************
'����:ɨ�븶�ؼ�
'����������:
'  һ�����ŷ���������
'       1.zlInit-��ʼ������(��Ҫ�Ǹ�Oracle���Ӽ�����ɨ�븶��������)
'       2.zlReReadQRCode:���¶�ȡ��ά�����
'  �����ڲ�����������
'       1.GetQRCodePayment-��ȡɨ�븶����
'       2.ReadQRCode-��ȡɨ�븶��Ҫ֧���Ķ�ά��
'       3.SetBorderVisible:���ñ߿��ߵ���ʾ
'       4.DrawButtonStyle�����ư�ť����ʽ
'       5.ClearButtonTag:�����ť������Ϣ
'       6.Refresh:����ˢ�½���
'����:
'   1.BorderStyle-�ؼ��Ƿ���߿���
'   2.Appearance-�ؼ�����ʾ��ʽ
'   3.CardTypeIDs-����֧�ֵĿ����IDs,����ö��ŷָ�,ֻ�����ԣ��ɷ�����zlInit����)
'   4.PicAlignment-ͼ����뷽ʽ(������ı�����)
'   5.CaptionAlignMent-�ı����뷽ʽ
'   6.Caption-��ť�ı�
'�¼�:
'   1.zlErrShow-������ʾ������������ʱ���������¼�)
'   2.zlQRCodePayment-��ȡ��ά��ɹ���ʧ�ܺ󣬷���֧���¼�
'   3.zlGetPayMoney-��ȡ����֧���Ľ��(�ڵ����ťʱ�������¼�)
'˵��:
'  1.��������Ҫʹ��ʱ������Ҫ��ϡ�zlQRCodePayMent.clsQRCodePayment������һ�����Ч
'  2.����˳��:
'      1)���ȣ����á�zlInit�����г�ʼ����δ��ʼ���ɹ���������ʹ�øÿؼ����������Բ���ʾ�ÿؼ�
'      2)��Σ�ͨ���¼���GetQrCodePayment�����ر���Ҫ֧���Ľ��
'      2)���,ͨ���¼�"zlQRCodePayment"�¼�����֧��
'      3)�������������ͨ����zlErrShow���¼���ʾ��صĴ�����Ϣ.
'      4)�����ͨ�����ʵ�֣������"zlReReadQRCode"����
'����:���˺�
'����:2019-03-04 19:19:10
'*********************************************************************************************************************************************

Public Enum PayButton_Appearance
    ShowFlat = 0
    Show3D = 1 '��ǳ�İ�ť
    ShowEdge3D = 2 '����İ�ť
End Enum
Public Enum PayButton_BorderStyle
    ShowNone = 0
    ShowFixed_Single = 1
End Enum

Public Enum PayButton_Alignment
    LeftAgnmt = 0
    CenterAgnmt
    RightAgnmt
End Enum

Public Enum PayButton_PicAlignment
    TxtLeftAgnmt = 0
    TxtDownCenterAgnmt = 1
    TxtTopCenterAgnmt = 2
    TxtRightAgnmt = 3
End Enum

'Const m_def_Enabled = 0
Const m_def_BorderStyle = 0
Const m_def_Appearance = 0

'���Ա���:
Dim m_Caption As String
Dim m_CaptionAlignMent As PayButton_Alignment
Dim m_PicAlignMent As PayButton_PicAlignment

Dim m_CardTypeIDs As String '����֧�ֵĿ����ID
'Dim m_Enabled As Boolean
Dim m_BorderStyle As PayButton_BorderStyle
Dim m_Appearance As PayButton_Appearance

'�¼�����:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Event zlErrShow(ByVal strErrMsg As String, ByVal lngErrNum As Long)
Event zlGetPayMoney(ByRef dblMoney As Double, ByRef strExpend As String, ByRef blnCancel As Boolean)
'blnCancel:���ʱ:true:��ʾ��ȡ��ά��ʧ�ܣ������ȡ��ά��ɹ�;����ʱ:��ʾ��ֹ����֧��
'lngCardTypeID:��ȡ��ά��ɹ�ʱ��Ϊ��ά��֧���Ŀ����ID,����Ϊ0
Event zlQRCodePayment(ByVal lngCardTypeID As Long, ByVal strPayMentQRCode As String, ByVal strExpendXML As String, ByRef blnCancel As Boolean)


'ģ�鼶����
Private mcnOracle As ADODB.Connection
Private mstrDBUser As String
Private mlngSys As Long, mlngModul As Long
Private mobjQRCodePayment As Object '��ά��֧����ɨ��֧������
Private mfrmMain As Object
'ȱʡ����ֵ:
Const m_def_Caption = ""
Const m_def_CaptionAlignMent = 0
Const m_def_PicAlignMent = 0
Const m_def_CardTypeIDs = ""

Public Function zlInit(ByVal frmMain As Object, ByVal strCardTypeIDs As String, Optional ByVal lngSys As Long, Optional ByVal lngModul As Long, _
    Optional cnOracle As ADODB.Connection, Optional strDBUser As String, Optional ByRef strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��
    '���:frmMain-���õ�������
    '     lngSys : ϵͳ���
    '     lngModul:��Ҫִ�еĹ������
    '     cnOracle:����������ݿ�����
    '     strCardTypeIDs-����֧�ֵ�ɨ�븶�����IDs(����ö��ŷָ�),��:1,2,3...
    '����:strErrMsg_out-���صĴ�����Ϣ
    '����:���˺�
    '����:2019-03-04 19:36:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objQRCode As Object
    
    Set mcnOracle = cnOracle: mstrDBUser = strDBUser: mlngModul = lngModul: mlngSys = lngSys
    Set mfrmMain = frmMain: m_CardTypeIDs = strCardTypeIDs
    
    If GetQrCodePayment(objQRCode, strErrMsg_out) = False Then Exit Function
    zlInit = True
    Exit Function
errHandle:
    strErrMsg_out = Err.Description
    RaiseEvent zlErrShow(strErrMsg_out, Err.Number)
End Function
Public Function zlReReadQRCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���¶�ȡ��ά�����
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-03-11 20:14:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlReReadQRCode = ReadQRCode
End Function
Private Function GetQrCodePayment(ByRef objQRCode_Out As Object, Optional ByRef strErrMsg_out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡɨ�븶����
    '����:objQRCode_Out-ɨ�븶����
    '     ErrMsg_out-��ȡɨ�븶����ʧ��
    '����:��ȡ�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-03-04 19:36:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnInitComponent As Boolean, strExpandXML As String
    If Not mobjQRCodePayment Is Nothing Then
        Set objQRCode_Out = mobjQRCodePayment
        GetQrCodePayment = True: Exit Function
    End If
    
    Err = 0: On Error Resume Next
    Set mobjQRCodePayment = CreateObject("zlReadQRCode.clsReadQRCode")    '�̶�����
    If Err.Number <> 0 Then
        'strErrMsg_out = "����ɨ�븶����(zlQRCodePayMent.clsQRCodePayment)ʧ��,����ö����Ƿ���ȷ��"
        'RaiseEvent zlErrShow(strErrMsg_out, 0)
        Exit Function
    End If
    Err = 0: On Error GoTo ErrHand:
    '��ʼ���ӿڲ���
    '����:zlInitComponents (��ʼ���ӿڲ���)
    '���:   lngSys-�����ϵͳ��
    '       strDBUser-���ݿ��û���
    '       cnOracle -HIS/��������
    '       strCardTypeIDs-����֧�ֵ�ɨ�븶�����IDs(����ö��ŷָ�),��:1,2,3...
    '       strExpandXML-��չ��Ϣ,����,���Ժ���չ
    blnInitComponent = mobjQRCodePayment.zlInitComponents(mlngSys, mstrDBUser, mcnOracle, m_CardTypeIDs, strExpandXML)
    If Not blnInitComponent Then
        Exit Function
    End If
    Set objQRCode_Out = mobjQRCodePayment
    GetQrCodePayment = blnInitComponent
    Exit Function
ErrHand:
    strErrMsg_out = Err.Description
    RaiseEvent zlErrShow(strErrMsg_out, Err.Number)
End Function

Private Function ReadQRCode() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡɨ�븶��Ҫ֧���Ķ�ά��
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-03-04 20:03:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objQRCode As Object, strErrMsg_out As String, lngCardTypeID_Out As Long, strQRCode_Out As String, strExpendXML As String
    Dim blnCancel As Boolean, dblMoney As Double, strExpend As String
    On Error GoTo errHandle
    
    '1.��ȡ��ǰɨ�븶��֧�����
    
    blnCancel = False: strExpendXML = "": dblMoney = 0
    RaiseEvent zlGetPayMoney(dblMoney, strExpend, blnCancel) 'strExpend:�����ô�
    If blnCancel Or dblMoney = 0 Then Exit Function
    
    If dblMoney = 0 Then
        strErrMsg_out = "������Ϊ�㣬����Ҫ����ɨ�븶��!"
        RaiseEvent zlErrShow(strErrMsg_out, 0)
        Exit Function
    End If
    
    If dblMoney < 0 Then
        strErrMsg_out = "��֧�ָ�����ɨ�븶!"
        RaiseEvent zlErrShow(strErrMsg_out, 0)
        Exit Function
    End If
    
    strErrMsg_out = ""
    If GetQrCodePayment(objQRCode, strErrMsg_out) = False Then Exit Function
    '���ö�ȡ֧����
    '    zlReadQRCode(frmMain As Object, _
    '    ByVal lngModule As Long,
    '    ByVal dblMoney As Double,
    '    ByRef lngCardTypeID_Out As  Long, _
    '    ByRef strQRCode As String,byref strExpand As String, _) As Boolean
    '    '----------------------------------------------------------------------------------------------------------------------------------------
    '    '����:��ȡ֧���Ķ�ά�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpendXML-��չ����,������
    '    '����:lngCardTypeID_Out-���ص�֧�������ID
    '    '       strQRCode_Out-���ص�֧����
    '    '       strExpendXML-���Ժ���չ
    '    '����:��������    True:���óɹ�,False:����ʧ��
    If objQRCode.ReadQRCode(mfrmMain, mlngModul, dblMoney, lngCardTypeID_Out, strQRCode_Out, strExpendXML) = False Then
        blnCancel = True
        RaiseEvent zlQRCodePayment(0, "", "", blnCancel)
        Exit Function
    End If
    RaiseEvent zlQRCodePayment(lngCardTypeID_Out, strQRCode_Out, strExpendXML, blnCancel)
    ReadQRCode = Not blnCancel
    Exit Function
errHandle:
    strErrMsg_out = Err.Description
    RaiseEvent zlErrShow(strErrMsg_out, Err.Number)
End Function

 
Private Sub DrawButtonStyle(ByVal intAppearance As PayButton_Appearance, Optional blnHotImg As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ư�ť����ʽ
    '���:0=ƽ��,-1=����,1=͹��(��ǳ�İ�ť),-2=���,2=��͹��(����İ�ť)
    '����:���˺�
    '����:2019-03-04 17:25:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intAlignMent As gAlignment
    Dim imgPicture As IPictureDisp
    Dim intStyle As Integer
    
    intAlignMent = m_CaptionAlignMent
    
    If blnHotImg Or m_Appearance <> ShowFlat And intAppearance <> -1 And UserControl.Enabled Then
        Set imgPicture = imgHot.Picture
    Else
        Set imgPicture = ImgNormal.Picture
    End If
    
    'intStyle=0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��
    'intType:0-����;1-OnlyDown;2-OnlyKind
    Select Case intAppearance
    Case ShowFlat 'ƽ��
        intStyle = 0
    Case ShowEdge3D '����İ�ť
        intStyle = 2
    Case Show3D '��ǳ�İ�ť
        intStyle = 1
    Case Else
        intStyle = intAppearance
    End Select
    zlRaisEffectEx picButton, intStyle, m_Caption, m_CaptionAlignMent, imgPicture, m_PicAlignMent
End Sub
Private Sub ClearButtonTag()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ť������Ϣ
    '����:���˺�
    '����:2019-03-04 17:40:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
     shapBack.BorderColor = &H8000000A
     picButton.Tag = ""
End Sub

Private Sub SetBorderVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ñ߿��ߵ���ʾ
    '���:
    '����:���˺�
    '����:2019-03-04 18:51:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    shapBack.Visible = BorderStyle = ShowFixed_Single And m_Appearance = ShowFlat
    Call UserControl_Resize
End Sub

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=5
Public Sub Refresh()
     Call SetBorderVisible
     Call DrawButtonStyle(m_Appearance)
End Sub


Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    shapBack.Visible = False
    Call DrawButtonStyle(-1)  '���ư�ť:0=ƽ��,-1=����,1=͹��(��ǳ�İ�ť),-2=���,2=��͹��(����İ�ť)
    Call ClearButtonTag   '�����ť������Ϣ
End Sub
Private Sub picButton_Resize()
    Err = 0: On Error Resume Next
    Call DrawButtonStyle(Appearance)    '���ư�ť
    Call ClearButtonTag   '�����ť������Ϣ
End Sub


Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Appearance = ShowEdge3D Or Appearance = Show3D Then Exit Sub
    'ֻ��ƽ��ģ����л���
    If picButton.Tag = "In" Then
         If X < 0 Or Y < 0 Or X > picButton.Width Or Y > picButton.Height Then
             picButton.Tag = ""
             ReleaseCapture
             shapBack.BorderColor = &H8000000A
             Call DrawButtonStyle(m_Appearance)  '���ư�ť
             Call SetBorderVisible
         End If
     Else
         picButton.Tag = "In"
         SetCapture picButton.hWnd
         shapBack.BorderColor = vbBlue
         Call DrawButtonStyle(IIf(shapBack.Visible, ShowFlat, Show3D), True)    '���ư�ť
     End If
End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    If Button <> 1 Then Exit Sub
    Call DrawButtonStyle(m_Appearance)  '���ư�ť
    Call ClearButtonTag
    Call SetBorderVisible   '��ʾ�߿���
    Call ReadQRCode
   ' RaiseEvent Click
End Sub
Private Sub UserControl_ExitFocus()
    'ƽ��,�ָ�ƽ��
    Call DrawButtonStyle(m_Appearance)  '���ư�ť
End Sub
'
'
''ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
''MemberInfo=0,0,0,0
'Public Property Get Enabled() As Boolean
'    Enabled = m_Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    m_Enabled = New_Enabled
'    PropertyChanged "Enabled"
'End Property
 
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As PayButton_BorderStyle
Attribute BorderStyle.VB_Description = "����/���ö���ı߿���ʽ��"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As PayButton_BorderStyle)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    Call SetBorderVisible
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=1,0,0,0
Public Property Get Appearance() As PayButton_Appearance
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As PayButton_Appearance)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    
    Call DrawButtonStyle(m_Appearance)  '���ư�ť
    Call ClearButtonTag
End Property
Private Sub UserControl_Resize()
    Dim lngX As Long, lngY As Long
    Err = 0: On Error Resume Next
    
    With shapBack
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    
    With UserControl
        lngX = 0: lngY = 0
        If shapBack.Visible Then
            lngX = shapBack.BorderWidth * Screen.TwipsPerPixelX
            lngY = shapBack.BorderWidth * Screen.TwipsPerPixelY
        End If
        picButton.Move .ScaleLeft + lngX, .ScaleTop + lngY, .ScaleWidth - (.ScaleLeft + lngX * 2), .ScaleHeight - (.ScaleTop + lngX * 2)
    End With
End Sub

Private Sub UserControl_Initialize()
    Err = 0: On Error Resume Next
    Call DrawButtonStyle(m_Appearance)  '���ư�ť
    Call ClearButtonTag
End Sub

'Ϊ�û��ؼ���ʼ������
Private Sub UserControl_InitProperties()
'    m_Enabled = m_def_Enabled
    m_BorderStyle = m_def_BorderStyle
    m_Appearance = m_def_Appearance
    Call SetBorderVisible
    m_CardTypeIDs = m_def_CardTypeIDs
    m_CaptionAlignMent = m_def_CaptionAlignMent
    m_PicAlignMent = m_def_PicAlignMent
    m_Caption = m_def_Caption
End Sub

'�Ӵ������м�������ֵ
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_CardTypeIDs = PropBag.ReadProperty("CardTypeIDs", m_def_CardTypeIDs)
    
    Call SetBorderVisible
    m_CaptionAlignMent = PropBag.ReadProperty("CaptionAlignMent", m_def_CaptionAlignMent)
    m_PicAlignMent = PropBag.ReadProperty("PicAlignMent", m_def_PicAlignMent)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Call DrawButtonStyle(m_Appearance)  '���ư�ť
    Call ClearButtonTag
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
'    picButton.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    picButton.ToolTipText = PropBag.ReadProperty("ToolTipString", "")
End Sub



Private Sub UserControl_Terminate()
    Err = 0: On Error Resume Next
    
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If Not mobjQRCodePayment Is Nothing Then Set mobjQRCodePayment = Nothing
    If Not mfrmMain Is Nothing Then Set mfrmMain = Nothing
End Sub

'������ֵд���洢��
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("CardTypeIDs", m_CardTypeIDs, m_def_CardTypeIDs)
    Call PropBag.WriteProperty("CaptionAlignMent", m_CaptionAlignMent, m_def_CaptionAlignMent)
    Call PropBag.WriteProperty("PicAlignMent", m_PicAlignMent, m_def_PicAlignMent)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
'    Call PropBag.WriteProperty("ToolTipText", picButton.ToolTipText, "")
    Call PropBag.WriteProperty("ToolTipString", picButton.ToolTipText, "")
End Sub
'
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
 
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=13,1,1,
Public Property Get CardTypeIDs() As String
    CardTypeIDs = m_CardTypeIDs
End Property

Public Property Let CardTypeIDs(ByVal New_CardTypeIDs As String)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_CardTypeIDs = New_CardTypeIDs
    PropertyChanged "CardTypeIDs"
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=1,0,0,0
Public Property Get CaptionAlignment() As PayButton_Alignment
    CaptionAlignment = m_CaptionAlignMent
End Property

Public Property Let CaptionAlignment(ByVal New_CaptionAlignment As PayButton_Alignment)
    m_CaptionAlignMent = New_CaptionAlignment
    PropertyChanged "CaptionAlignMent"
    Call DrawButtonStyle(m_Appearance)  '���ư�ť
    Call ClearButtonTag
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=1,0,0,0
Public Property Get PicAlignMent() As PayButton_PicAlignment
    PicAlignMent = m_PicAlignMent
End Property

Public Property Let PicAlignMent(ByVal New_PicAlignMent As PayButton_PicAlignment)
    m_PicAlignMent = New_PicAlignMent
    PropertyChanged "PicAlignMent"
    
    Call DrawButtonStyle(m_Appearance)  '���ư�ť
    Call ClearButtonTag
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MemberInfo=13,0,0,
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Call DrawButtonStyle(m_Appearance)  '���ư�ť
    Call ClearButtonTag
End Property

'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "����/����һ��ֵ������һ�������Ƿ���Ӧ�û������¼���"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
    Call DrawButtonStyle(m_Appearance)  '���ư�ť
    Call ClearButtonTag
End Property
 
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=picButton,picButton,-1,ToolTipText
Public Property Get ToolTipString() As String
Attribute ToolTipString.VB_Description = "����/���õ�����ڿؼ�����ͣʱ��ʾ���ı���"
    ToolTipString = picButton.ToolTipText
End Property
Public Property Let ToolTipString(ByVal New_ToolTipString As String)
    picButton.ToolTipText() = New_ToolTipString
    PropertyChanged "ToolTipString"
End Property

