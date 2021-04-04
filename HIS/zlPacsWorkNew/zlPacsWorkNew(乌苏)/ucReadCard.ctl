VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucReadCard 
   BackStyle       =   0  '͸��
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   ScaleHeight     =   975
   ScaleWidth      =   3285
   ToolboxBitmap   =   "ucReadCard.ctx":0000
   Begin VB.CommandButton cmdRead 
      Height          =   330
      Left            =   2890
      Picture         =   "ucReadCard.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   1305
      TabIndex        =   2
      Top             =   0
      Width           =   1335
      Begin VB.PictureBox picTag 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   60
         Picture         =   "ucReadCard.ctx":069C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   3
         Top             =   30
         Width           =   240
      End
      Begin VB.Label labCardType 
         AutoSize        =   -1  'True
         Caption         =   "        "
         Height          =   180
         Left            =   360
         TabIndex        =   4
         Top             =   45
         Width           =   600
      End
   End
   Begin MSComctlLib.ImageList imgCardType 
      Left            =   0
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucReadCard.ctx":0A26
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCardContext 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar tbrDown 
      Height          =   330
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   582
      ButtonWidth     =   1032
      ButtonHeight    =   529
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgCardType"
      DisabledImageList=   "imgCardType"
      HotImageList    =   "imgCardType"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbn_Select"
            Style           =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "ucReadCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const M_STR_CARD_SPLIT_CHA As String = ";"


Private mobjSquareCard As Object
Private mstrCardNames As String
Private mblnShowReadButton As Boolean
Private mblnAutoSize As Boolean

Private WithEvents mobjIdCard As clsIDcard
Attribute mobjIdCard.VB_VarHelpID = -1
Private WithEvents mobjIcCard As clsICCard
Attribute mobjIcCard.VB_VarHelpID = -1

Private mstrCurCardName As String
Private mlngCurKindId As Long               '��ǰ�����ID
Private mlngCurSwipingCardType As Long      '��ǰˢ������
Private mlngCardLen As Long                 '��ǰˢ������
Private mblnIsPwdInput  As Boolean          '��ǰ���Ƿ���Ҫ����¼��


Private maryKinds() As String   '���濨��Ϣ
Private mlngModule As Long

'������ˢ������¼��ɹ��󴥷����¼�
Public Event OnRead(ByVal strCardName As String, ByVal strFilter As String, ByVal strPatientId As String)
Public Event OnCardChange(ByVal strCardName As String)
      
Public Event OnResize()
                    

'����ͼƬ
Property Get Picture() As IPictureDisp
    Set Picture = picTag.Picture
End Property

Property Set Picture(value As IPictureDisp)
    Set picTag.Picture = value
    
    picTag.Visible = IIf(picTag.Picture = 0, False, True)
    
    Call AutoAdjustWidth
End Property


'�����ƣ��࿨֮���÷ֺţ���;�������
Property Get CardNames() As String
    CardNames = mstrCardNames
End Property

Property Let CardNames(value As String)
    mstrCardNames = value
    
    Call ConfigCardFace(value)
End Property


'�Զ���ʾ������ť
Property Get ShowReadButton() As Boolean
    ShowReadButton = mblnShowReadButton
End Property


Property Let ShowReadButton(value As Boolean)
    mblnShowReadButton = value
    cmdRead.Visible = value
End Property


'�Զ���С
Property Get AutoSize() As Boolean
    AutoSize = mblnAutoSize
End Property

Property Let AutoSize(value As Boolean)
    mblnAutoSize = value
    
    Call AutoAdjustWidth
End Property


'ˢ���ı�
Property Get CardText() As String
    CardText = txtCardContext.Text
End Property

Property Let CardText(value As String)
    txtCardContext.Text = value
End Property


'���õ�ǰˢ������
Property Get CurCardName() As String
    CurCardName = mstrCurCardName
End Property


Property Let CurCardName(value As String)
    mstrCurCardName = value
    
    Call ConfigCardFace(mstrCurCardName)
End Property


'�ؼ����
Property Get Handle() As Long
    Handle = UserControl.hWnd
End Property



Public Sub InitCardType(ByVal lngSys As Long, ByVal lngModule As Long, _
    ByVal strUser As String, cnOracle As ADODB.Connection)
    
    Dim aryKindInfo() As String
    Dim strKinds As String
    Dim i As Integer
    Dim bmCur As ButtonMenu
    Dim strFirstCard As String

    mlngModule = lngModule
    strFirstCard = ""
    
    '��ʼ�������㲿��
    Call mobjSquareCard.zlInitComponents(Me, lngModule, lngSys, strUser, cnOracle)

    aryKindInfo = Split(mstrCardNames, M_STR_CARD_SPLIT_CHA)
    
    strKinds = ""
    For i = 0 To UBound(aryKindInfo)
        If strKinds <> "" Then strKinds = strKinds & M_STR_CARD_SPLIT_CHA
        strKinds = strKinds & aryKindInfo(i) & "|" & aryKindInfo(i) & "|-1"
    Next i
    
    strKinds = strKinds & M_STR_CARD_SPLIT_CHA

    '��ȡ�ſ������Ϣ
    maryKinds = Split(mobjSquareCard.zlGetIDKindStr(strKinds), M_STR_CARD_SPLIT_CHA)
        
    '�������
    Call tbrDown.Buttons(1).ButtonMenus.Clear
    For i = 0 To UBound(maryKinds)
        aryKindInfo = Split(maryKinds(i), "|")
        If Trim(aryKindInfo(1)) <> "" Then
            Set bmCur = tbrDown.Buttons(1).ButtonMenus.Add()
            bmCur.Key = "tbm_" & i
            bmCur.Text = IIf(aryKindInfo(1) = "����", "��   ��", IIf(aryKindInfo(1) = "���֤��", "���֤", aryKindInfo(1))) & "(&" & IIf(i >= 9, Chr(65 + i - 9), i + 1) & ")"
            bmCur.Tag = aryKindInfo(1)
            
            If strFirstCard = "" Then strFirstCard = aryKindInfo(1)
        End If
    Next i
    
    '����ˢ��������ʾ
    Call ConfigCardFace(strFirstCard)
End Sub


Public Sub GetCardValue(ByRef strCardName As String, ByRef strCardText As String, ByRef lngPatientID As Long)
'��ȡˢ����ֵ������ж�Ӧ�Ŀ����ͣ��򷵻ز���ID,���򷵻�ԭֵ
 
    lngPatientID = 0
    strCardName = mstrCurCardName
    strCardText = txtCardContext.Text
    
    If mlngCurKindId > 0 Then
        Call mobjSquareCard.zlGetPatiID(IIf(mlngCurKindId > 0, mlngCurKindId, mstrCurCardName), strCardText, , lngPatientID)
    End If
End Sub


Private Function GetIDKindInfo(ByVal strKind As String) As String
'��ȡָ������Ϣ
On Error Resume Next
    Dim i As Long
    
    GetIDKindInfo = ""
    For i = 0 To UBound(maryKinds)
        If Trim(Split(maryKinds(i), "|")(1)) = strKind Then
            GetIDKindInfo = maryKinds(i)
            Exit Function
        End If
    Next i
End Function


Private Sub ConfigCardFace(ByVal strCardName As String)
'���ö�������
    Dim strCurKindInfo As String
    Dim aryKindInfo() As String
    
    mlngCurSwipingCardType = -1
    mlngCurKindId = 0
    mlngCardLen = 0
    mblnIsPwdInput = False
    mstrCurCardName = ""
    
    
    txtCardContext.Text = ""
    cmdRead.Visible = False
    
    If strCardName = "" Then Exit Sub
    
    strCurKindInfo = GetIDKindInfo(strCardName)
    
    If Trim(strCurKindInfo) <> "" Then
        aryKindInfo = Split(strCurKindInfo, "|")
        
        mlngCurKindId = Val(aryKindInfo(3))     '�����ID
        mlngCardLen = Val(aryKindInfo(4))    '���ų���
        mlngCurSwipingCardType = Val(aryKindInfo(2))   'ˢ������
        mblnIsPwdInput = IIf(Val(aryKindInfo(7)) = 0, False, True) '�Ƿ�������ʾ
    End If
    
    If mlngCurSwipingCardType = 1 Then '��ʾ����
        cmdRead.Visible = mblnShowReadButton
    Else
        cmdRead.Visible = False
    End If
    
    Call UserControl_Resize
    
    mstrCurCardName = strCardName
    
    labCardType.Caption = mstrCurCardName
    txtCardContext.PasswordChar = IIf(mblnIsPwdInput, "*", "")
    
    Call AutoAdjustWidth
End Sub



Private Sub cmdRead_Click()
'�����������
On Error GoTo ErrHandle

    Call StartReadCard
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Sub StartReadCard()
'��ʼ����
    Dim lngPatientID As Long
    
    If mlngCurSwipingCardType = 1 Then
        txtCardContext.Text = ReadCard(lngPatientID)

        RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, IIf(lngPatientID > 0, lngPatientID, ""))
    Else
        If mlngCurKindId > 0 Then
            Call mobjSquareCard.zlGetPatiID(mlngCurKindId, txtCardContext.Text, , lngPatientID)
            
            RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, IIf(lngPatientID > 0, lngPatientID, ""))
        Else
            RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, "")
        End If
    End If
    
    Call SelText
End Sub


Public Function ReadCard(ByRef lngPatientID As Long) As String
'ִ�ж�������
    Dim strExpand As String, strOutCardNO As String, strOutPatiInfoXML As String
    
    lngPatientID = 0
    ReadCard = ""
    
    If mlngCurSwipingCardType <> 1 Then Exit Function 'ˢ������Ϊ1��ʾ����
    
    strOutCardNO = ""
    
    If mlngCurKindId <> 0 Then
        '��ʼ����
        If mobjSquareCard.zlReadCard(Me, mlngModule, mlngCurKindId, True, strExpand, strOutCardNO, strOutPatiInfoXML) = False Then
            Exit Function
        End If
                
        ReadCard = strOutCardNO
        
        '�����ɹ��󣬸��ݶ�ȡ�������ݲ���
        If Not mobjSquareCard.zlGetPatiID(IIf(mlngCurKindId > 0, mlngCurKindId, mstrCurCardName), strOutCardNO, , lngPatientID) Then Exit Function
    End If

End Function


Private Sub labCardType_Click()
    Call picBack_Click
End Sub

Private Sub labCardType_DblClick()
    Call picBack_DblClick
End Sub

Private Sub mobjIcCard_ShowICCardInfo(ByVal strCardNo As String)
'��ȡIC��
On Error GoTo ErrHandle
    Dim strFilter As String
    Dim lngPatientID As Long
    
    txtCardContext.Text = strCardNo
    strFilter = strCardNo
    
    Call mobjIcCard.SetEnabled(False)

    If Not mobjSquareCard.zlGetPatiID("IC��", strCardNo, True, lngPatientID) Then Exit Sub

    RaiseEvent OnRead(mstrCurCardName, strFilter, IIf(lngPatientID > 0, lngPatientID, ""))
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub mobjIdCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
'��ȡ���֤
On Error GoTo ErrHandle
    txtCardContext.Text = strID
    
    Call mobjIdCard.SetEnabled(False)
    
    RaiseEvent OnRead(mstrCurCardName, strID, "")
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub picBack_Click()
'�����¼�
On Error GoTo ErrHandle
    Call StartReadCard
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub picBack_DblClick()
'���˫���¼�
On Error GoTo ErrHandle
    Call StartReadCard
    
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub picTag_Click()
    Call picBack_Click
End Sub

Private Sub picTag_DblClick()
    Call picBack_DblClick
End Sub

Private Sub tbrDown_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'����ѡ��Ŀ�����
On Error GoTo ErrHandle
    Call ConfigCardFace(Mid(ButtonMenu.Text, 1, InStr(ButtonMenu.Text, "(") - 1))
    
    Call AutoAdjustWidth
    
    RaiseEvent OnCardChange(mstrCurCardName)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub txtCardContext_Change()
On Error GoTo ErrHandle
    If Trim(txtCardContext.Text) = "" Then
        Call mobjIdCard.SetEnabled(True)
        Call mobjIcCard.SetEnabled(True)
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtCardContext_DblClick()
'    Call picBack_DblClick
End Sub


Private Sub txtCardContext_GotFocus()
'��������ƶ����ÿؼ��ϣ���ȫѡ�ı�
On Error Resume Next
    
    Call mobjIdCard.SetEnabled(True)
    Call mobjIcCard.SetEnabled(True)
    
    If txtCardContext.Text <> "" Then Call zlControl.TxtSelAll(txtCardContext)
err.Clear
End Sub


Private Sub txtCardContext_KeyPress(KeyAscii As Integer)
'¼���¼�
On Error GoTo ErrHandle
    Dim blnCard As Boolean
    Dim lngPatientID As Long
        
    If KeyAscii = 13 Then
        If mlngCurSwipingCardType = 1 Then
            txtCardContext.Text = ReadCard(lngPatientID)
            RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, IIf(lngPatientID > 0, lngPatientID, ""))
        Else
            If mlngCurKindId > 0 Then
                Call mobjSquareCard.zlGetPatiID(mlngCurKindId, txtCardContext.Text, , lngPatientID)
                
                RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, IIf(lngPatientID > 0, lngPatientID, ""))
            Else
                RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, "")
            End If
            
        End If
        
        Call SelText
        
        Exit Sub
    End If
    
    If mlngCurSwipingCardType = 0 Then  '����ˢ������
        If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        
        blnCard = zlCommFun.InputIsCard(txtCardContext, KeyAscii, mblnIsPwdInput)
        If blnCard And Len(txtCardContext.Text) = mlngCardLen - 1 And KeyAscii <> 8 Then  'ˢ����ϴ���
        
            txtCardContext.Text = txtCardContext.Text & Chr(KeyAscii)
            txtCardContext.SelStart = Len(txtCardContext.Text)
            
            KeyAscii = 0
            
            Call zlControl.TxtSelAll(txtCardContext)
            
            If mlngCurKindId > 0 Then
                Call mobjSquareCard.zlGetPatiID(mlngCurKindId, txtCardContext.Text, , lngPatientID)
                
                RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, IIf(lngPatientID > 0, lngPatientID, ""))
            Else
                RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, "")
            End If
        End If
    End If
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtCardContext_LostFocus()
On Error Resume Next
    Call mobjIdCard.SetEnabled(False)
    Call mobjIcCard.SetEnabled(False)
    
    err.Clear
End Sub

Private Sub txtCardContext_Validate(Cancel As Boolean)
'���벿�ֵ��ݺţ�����ȫ�����ݺ�
On Error Resume Next
    If InStr(mstrCurCardName, "���ݺ�") > 0 Then
        If IsNumeric(txtCardContext.Text) Then
            txtCardContext.Text = GetFullNO(txtCardContext.Text, 0)
        End If
    End If
err.Clear
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
    '���������㲿��
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    
    Set mobjIdCard = New zlIDCard.clsIDcard
    Set mobjIcCard = New zlICCard.clsICCard
    
    Call mobjIdCard.SetEnabled(False)
    Call mobjIcCard.SetEnabled(False)
    
    Call ConfigCardFace("")
err.Clear
End Sub


Public Sub SelText()
'ѡ���ı���
    Call zlControl.TxtSelAll(txtCardContext)
End Sub


Private Sub UserControl_Paint()
'    If Not UserControl.Enabled Then
'        txtCardContext.BackColor = UserControl.BackColor
'    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'��ȡ�������
    mstrCardNames = PropBag.ReadProperty("CardNames", "")
    mblnShowReadButton = PropBag.ReadProperty("ShowReadButton", True)
    mblnAutoSize = PropBag.ReadProperty("AutoSize", False)
    Set picTag.Picture = PropBag.ReadProperty("Picture", Nothing)
    picTag.Visible = IIf(picTag.Picture = 0, False, True)
    
    Call AutoAdjustWidth
End Sub


Private Sub AutoAdjustWidth()
'�Զ�����������
    Dim lngLabInc As Long

    
    If mblnAutoSize Then
        picBack.Width = labCardType.Width + picTag.Width + 180
    Else
        picBack.Width = 1075
    End If
    
'    Extender.Width = picBack.Width + 310 + txtCardContext.Width + IIf(cmdRead.Visible, cmdRead.Width, 0)
    
    Call UserControl_Resize
End Sub


Private Sub UserControl_Resize()
'���Ʋ�����С
On Error Resume Next
    Extender.Height = txtCardContext.Height
    
    tbrDown.Left = picBack.Left + picBack.Width - tbrDown.Width + 310
    txtCardContext.Left = tbrDown.Left + tbrDown.Width - 20
    txtCardContext.Width = Extender.Width - picBack.Width - 310 - IIf(cmdRead.Visible, cmdRead.Width + 10, 0)
    
    If picTag.Picture <> 0 Then
        labCardType.Left = picTag.Left + picTag.Width + 30
    Else
        labCardType.Left = 30
    End If

    cmdRead.Left = txtCardContext.Left + txtCardContext.Width
    
    RaiseEvent OnResize
err.Clear
End Sub


Private Sub UserControl_Terminate()
'�ͷŲ����������Ķ���
On Error Resume Next
    Set mobjSquareCard = Nothing
    
    Call mobjIdCard.SetEnabled(False)
    Call mobjIcCard.SetEnabled(False)
    
    Set mobjIdCard = Nothing
    Set mobjIcCard = Nothing
err.Clear
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'д������
    Call PropBag.WriteProperty("CardNames", mstrCardNames, "")
    Call PropBag.WriteProperty("ShowReadButton", mblnShowReadButton, True)
    Call PropBag.WriteProperty("AutoSize", mblnAutoSize, False)
    Call PropBag.WriteProperty("Picture", picTag.Picture, Nothing)
End Sub
