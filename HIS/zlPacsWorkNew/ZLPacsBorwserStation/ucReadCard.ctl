VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucReadCard 
   BackStyle       =   0  '透明
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3870
   ScaleHeight     =   975
   ScaleWidth      =   3870
   ToolboxBitmap   =   "ucReadCard.ctx":0000
   Begin VB.CommandButton cmdRead 
      Height          =   330
      Left            =   3480
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
      ScaleWidth      =   1545
      TabIndex        =   2
      Top             =   0
      Width           =   1575
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
         Width           =   720
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
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin MSComctlLib.Toolbar tbrDown 
      Height          =   330
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   880
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

Private mstrCurCardName As String
Private mlngCurKindId As Long               '当前卡类别ID
Private mlngCurSwipingCardType As Long      '当前刷卡类型
Private mlngCardLen As Long                 '当前刷卡长度
Private mblnIsPwdInput  As Boolean          '当前卡是否需要隐藏录入


Private maryKinds() As String   '保存卡信息
Private mlngModule As Long

'读卡，刷卡或者录入成功后触发的事件
Public Event OnRead(ByVal strCardName As String, ByVal strCardText As String, ByVal lngPatientID As Long)

Public Event OnKeyPress(ByRef KeyAscii As Integer)

Public Event OnClick(ByVal strCardName As String, ByVal strCardText As String, _
                    ByVal lngKindId As Long, ByVal lngCardLen As Long, ByVal lngSwipingType As Long, ByVal blnIsPwdInput As Boolean)
                    
Public Event OnDblClick(ByVal strCardName As String, ByVal strCardText As String, _
                    ByVal lngKindId As Long, ByVal lngCardLen As Long, ByVal lngSwipingType As Long, ByVal blnIsPwdInput As Boolean)
                    
Public Event OnResize()
                    

'读卡图片
Property Get Picture() As IPictureDisp
    Set Picture = picTag.Picture
End Property

Property Set Picture(value As IPictureDisp)
    Set picTag.Picture = value
End Property


'卡名称，多卡之间用分号（“;”）间隔
Property Get CardNames() As String
    CardNames = mstrCardNames
End Property

Property Let CardNames(value As String)
    mstrCardNames = value
    
    Call ConfigCardFace(value)
End Property


'自动显示读卡按钮
Property Get ShowReadButton() As Boolean
    ShowReadButton = mblnShowReadButton
End Property


Property Let ShowReadButton(value As Boolean)
    mblnShowReadButton = value
    cmdRead.Visible = value
End Property


'自动大小
Property Get AutoSize() As Boolean
    AutoSize = mblnAutoSize
End Property

Property Let AutoSize(value As Boolean)
    mblnAutoSize = value
    
    Call AutoAdjustWidth
End Property


'刷卡文本
Property Get CardText() As String
    CardText = txtCardContext.Text
End Property

Property Let CardText(value As String)
    txtCardContext.Text = value
End Property


'配置当前刷卡类型
Property Get CurCardName() As String
    CurCardName = mstrCurCardName
End Property


Property Let CurCardName(value As String)
    mstrCurCardName = value
    
    Call ConfigCardFace(mstrCurCardName)
End Property


'控件句柄
Property Get Handle() As Long
    Handle = UserControl.Hwnd
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
    
    '初始化卡结算部件
    Call mobjSquareCard.zlInitComponents(Me, lngModule, lngSys, strUser, cnOracle)

    aryKindInfo = Split(mstrCardNames, M_STR_CARD_SPLIT_CHA)
    
    strKinds = ""
    For i = 0 To UBound(aryKindInfo)
        If strKinds <> "" Then strKinds = strKinds & M_STR_CARD_SPLIT_CHA
        strKinds = strKinds & aryKindInfo(i) & "|" & aryKindInfo(i) & "|-1"
    Next i
    
    strKinds = strKinds & M_STR_CARD_SPLIT_CHA

    '获取磁卡类别信息
    maryKinds = Split(mobjSquareCard.zlGetIDKindStr(strKinds), M_STR_CARD_SPLIT_CHA)
        
    '加载类别
    Call tbrDown.Buttons(1).ButtonMenus.Clear
    For i = 0 To UBound(maryKinds)
        aryKindInfo = Split(maryKinds(i), "|")
        If Trim(aryKindInfo(1)) <> "" Then
            Set bmCur = tbrDown.Buttons(1).ButtonMenus.Add()
            bmCur.Key = "tbm_" & i
            bmCur.Text = IIf(aryKindInfo(1) = "姓名", "姓   名", IIf(aryKindInfo(1) = "身份证号", "身份证", aryKindInfo(1))) & "(&" & IIf(i >= 9, Chr(65 + i - 9), i + 1) & ")"
            
            If strFirstCard = "" Then strFirstCard = aryKindInfo(1)
        End If
    Next i
    
    '配置刷卡界面显示
    Call ConfigCardFace(strFirstCard)
End Sub


Public Sub GetCardValue(ByRef strCardName As String, ByRef strCardText As String, ByRef lngPatientID As Long)
'获取刷卡的值，如果有对应的卡类型，则返回病人ID,否则返回原值
 
    lngPatientID = 0
    strCardName = mstrCurCardName
    strCardText = txtCardContext.Text
    
    If mlngCurKindId > 0 Then
        Call mobjSquareCard.zlGetPatiID(IIf(mlngCurKindId > 0, mlngCurKindId, mstrCurCardName), strCardText, , lngPatientID)
    End If
End Sub


Private Function GetIDKindInfo(ByVal strKind As String) As String
'获取指定卡信息
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
'配置读卡界面
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
        
        mlngCurKindId = Val(aryKindInfo(3))     '卡类别ID
        mlngCardLen = Val(aryKindInfo(4))    '卡号长度
        mlngCurSwipingCardType = Val(aryKindInfo(2))   '刷卡类型
        mblnIsPwdInput = IIf(Val(aryKindInfo(7)) = 0, False, True) '是否密文显示
    End If
    
    If mlngCurSwipingCardType = 1 Then '表示读卡
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
'处理读卡操作
On Error GoTo errHandle
    Dim lngPatientID As Long
    
    txtCardContext.Text = ReadCard(lngPatientID)

    RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, lngPatientID)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Public Function ReadCard(ByRef lngPatientID As Long) As String
'执行读卡操作
    Dim strExpand As String, strOutCardNO As String, strOutPatiInfoXML As String
    
    lngPatientID = 0
    ReadCard = ""
    
    If mlngCurSwipingCardType <> 1 Then Exit Function '刷卡类型为1表示读卡
    
    strOutCardNO = ""
    
    If mlngCurKindId <> 0 Then
        '开始读卡
        If mobjSquareCard.zlReadCard(Me, mlngModule, mlngCurKindId, True, strExpand, strOutCardNO, strOutPatiInfoXML) = False Then
            Exit Function
        End If
                
        ReadCard = strOutCardNO
        
        '读卡成功后，根据读取来的数据查找
        If Not mobjSquareCard.zlGetPatiID(IIf(mlngCurKindId > 0, mlngCurKindId, mstrCurCardName), strOutCardNO, , lngPatientID) Then Exit Function
    End If

End Function


Private Sub labCardType_Click()
    Call picBack_Click
End Sub

Private Sub labCardType_DblClick()
    Call picBack_DblClick
End Sub

Private Sub picBack_Click()
'单击事件
On Error GoTo errHandle
    RaiseEvent OnClick(mstrCurCardName, txtCardContext.Text, mlngCurKindId, mlngCardLen, mlngCurSwipingCardType, mblnIsPwdInput)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub picBack_DblClick()
'鼠标双击事件
On Error GoTo errHandle
    RaiseEvent OnDblClick(mstrCurCardName, txtCardContext.Text, mlngCurKindId, mlngCardLen, mlngCurSwipingCardType, mblnIsPwdInput)
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub picTag_Click()
    Call picBack_Click
End Sub

Private Sub picTag_DblClick()
    Call picBack_DblClick
End Sub

Private Sub tbrDown_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'配置选择的卡类型
On Error GoTo errHandle
    Call ConfigCardFace(Mid(ButtonMenu.Text, 1, InStr(ButtonMenu.Text, "(") - 1))
    
    Call AutoAdjustWidth
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub txtCardContext_DblClick()
    Call picBack_DblClick
End Sub


Private Sub txtCardContext_GotFocus()
'如果焦点移动到该控件上，则全选文本
On Error Resume Next
    If txtCardContext.Text <> "" Then Call zlControl.TxtSelAll(txtCardContext)
err.Clear
End Sub


Private Sub txtCardContext_KeyPress(KeyAscii As Integer)
'录入事件
On Error GoTo errHandle
    Dim blnCard As Boolean
    Dim lngPatientID As Long
    
    RaiseEvent OnKeyPress(KeyAscii)
    
    If KeyAscii = 13 Then
        If mlngCurSwipingCardType = 1 Then
            txtCardContext.Text = ReadCard(lngPatientID)
            RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, lngPatientID)
        Else
            If mlngCurKindId > 0 Then
                Call mobjSquareCard.zlGetPatiID(IIf(mlngCurKindId > 0, mlngCurKindId, mstrCurCardName), txtCardContext.Text, , lngPatientID)
            End If
            
            RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, lngPatientID)
        End If
        
        Exit Sub
    End If
    
    If mlngCurSwipingCardType = 0 Then  '处理刷卡操作
        If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
        
        blnCard = zlCommFun.InputIsCard(txtCardContext, KeyAscii, mblnIsPwdInput)
        If blnCard And Len(txtCardContext.Text) = mlngCardLen - 1 And KeyAscii <> 8 Then  '刷卡完毕处理
        
            txtCardContext.Text = txtCardContext.Text & Chr(KeyAscii)
            txtCardContext.SelStart = Len(txtCardContext.Text)
            
            KeyAscii = 0
            
            Call zlControl.TxtSelAll(txtCardContext)
            
            If mlngCurKindId > 0 Then
                Call mobjSquareCard.zlGetPatiID(IIf(mlngCurKindId > 0, mlngCurKindId, mstrCurCardName), txtCardContext.Text, , lngPatientID)
            End If
    
            RaiseEvent OnRead(mstrCurCardName, txtCardContext.Text, lngPatientID)
        End If
    End If
Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub txtCardContext_Validate(Cancel As Boolean)
'输入部分单据号，返回全部单据号
On Error Resume Next
    If InStr(mstrCurCardName, "单据号") > 0 Then
        If IsNumeric(txtCardContext.Text) Then
            txtCardContext.Text = GetFullNO(txtCardContext.Text, 0)
        End If
    End If
err.Clear
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
    '创建卡结算部件
    Set mobjSquareCard = CreateObject("zl9CardSquare.clsCardSquare")
    
    Call ConfigCardFace("")
err.Clear
End Sub


Public Sub SelText()
'选中文本行
    Call zlControl.TxtSelAll(txtCardContext)
End Sub


Private Sub UserControl_Paint()
'    If Not UserControl.Enabled Then
'        txtCardContext.BackColor = UserControl.BackColor
'    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'读取组件属性
    mstrCardNames = PropBag.ReadProperty("CardNames", "")
    mblnShowReadButton = PropBag.ReadProperty("ShowReadButton", True)
    mblnAutoSize = PropBag.ReadProperty("AutoSize", False)
    Set picTag.Picture = PropBag.ReadProperty("Picture", Nothing)
    
    Call AutoAdjustWidth
End Sub


Private Sub AutoAdjustWidth()
'自动调节组件宽度
    Dim lngLabInc As Long

    
    If mblnAutoSize Then
        picBack.Width = labCardType.Width + picTag.Width + 180
    Else
        picBack.Width = 1575
    End If
    
    Extender.Width = picBack.Width + 310 + txtCardContext.Width + IIf(cmdRead.Visible, cmdRead.Width, 0)
    
    Call UserControl_Resize
End Sub


Private Sub UserControl_Resize()
'控制部件大小
On Error Resume Next
    Extender.Height = txtCardContext.Height
    
    tbrDown.Left = picBack.Left + picBack.Width - tbrDown.Width + 310
    txtCardContext.Left = tbrDown.Left + tbrDown.Width - 20
    txtCardContext.Width = Extender.Width - picBack.Width - 310 - IIf(cmdRead.Visible, cmdRead.Width + 10, 0)

    cmdRead.Left = txtCardContext.Left + txtCardContext.Width
    
    RaiseEvent OnResize
err.Clear
End Sub


Private Sub UserControl_Terminate()
'释放部件所创建的对象
On Error Resume Next
    Set mobjSquareCard = Nothing
err.Clear
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'写入属性
    Call PropBag.WriteProperty("CardNames", mstrCardNames, "")
    Call PropBag.WriteProperty("ShowReadButton", mblnShowReadButton, True)
    Call PropBag.WriteProperty("AutoSize", mblnAutoSize, False)
    Call PropBag.WriteProperty("Picture", picTag.Picture, Nothing)
End Sub
