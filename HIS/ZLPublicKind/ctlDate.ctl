VERSION 5.00
Begin VB.UserControl ctlDate 
   BackColor       =   &H80000005&
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1590
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   510
   ScaleWidth      =   1590
   Begin VB.PictureBox picDate 
      BackColor       =   &H80000005&
      Height          =   300
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   1380
      TabIndex        =   0
      Top             =   30
      Width           =   1440
      Begin VB.CommandButton cmdDate 
         Caption         =   "…"
         Height          =   285
         Left            =   1155
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
      End
      Begin VB.TextBox txtEdit 
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   915
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "03"
         Top             =   52
         Width           =   195
      End
      Begin VB.TextBox txtEdit 
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   555
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "02"
         Top             =   52
         Width           =   180
      End
      Begin VB.TextBox txtEdit 
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   30
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "2010"
         Top             =   52
         Width           =   360
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblDaySplit 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   180
         Left            =   795
         TabIndex        =   5
         Top             =   60
         Width           =   90
      End
      Begin VB.Label lblMonthSplit 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         Height          =   180
         Left            =   435
         TabIndex        =   4
         Top             =   60
         Width           =   90
      End
   End
End
Attribute VB_Name = "ctlDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Enum TxtIndex
    Idx_txtYear = 0
    Idx_txtMonth = 1
    Idx_txtDay = 2
End Enum
Public Event CmdDownClick()
Public Event Change()
Public Event LastDayInput()
Private mstrCustomFormat As String '自定义格式串
Private mblnNotChange As Boolean
'缺省属性值:
Private m_def_Value   As Date
Private m_def_MaxDate As Date
Private m_def_MinDate  As Date
Const m_def_CustomFormat = "YYYY-MM-DD"
'属性变量:
Dim m_Value As Date
Dim m_MaxDate As Date
Dim m_MinDate As Date
Dim m_CustomFormat As String
Private mintLastSetFocus As Integer
  
Private Sub cmdDate_Click()
        RaiseEvent CmdDownClick
End Sub

Private Sub picDate_GotFocus()
    Call UserControl_GotFocus
End Sub

Private Sub picDate_Resize()
    Err = 0: On Error Resume Next
    With picDate
         cmdDate.Left = picDate.ScaleLeft + picDate.ScaleWidth - cmdDate.Width
    End With
End Sub

Private Sub txtEdit_Change(Index As Integer)
    Dim blnChange As Boolean
    If mblnNotChange = True Then Exit Sub
   Select Case Index
    Case Idx_txtYear  '年
            If txtEdit(Index).SelStart >= 4 Then
                 If VeryYear = True Then GoTo GoFocus:
                 RaiseEvent Change
                 SendKeys "{tab}"
            End If
    Case Idx_txtMonth '月
            '先看检证:
            If txtEdit(Index).SelStart >= 2 Then
                '先看检证:
                If VeryMonth = True Then GoTo GoFocus:
                RaiseEvent Change
                SendKeys "{tab}"
            End If
    Case Idx_txtDay '日
            '先看检证:
            If txtEdit(Index).SelStart >= 2 Then
                '先看检证:
                If VeryDay = True Then GoTo GoFocus:
                RaiseEvent Change
                RaiseEvent LastDayInput
            End If
    End Select
    Exit Sub
GoFocus:
    TxtSelAll txtEdit(Index)
End Sub
Private Function GetLastDateOfMonth() As Integer
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取指定月份的最后一天
    '编制：刘兴洪
    '日期：2010-06-23 17:45:27
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim dtTemp1 As Date, dtTemp2 As Date
    Dim strTemp As String
    strTemp = txtEdit(Idx_txtYear) & "-" & txtEdit(Idx_txtMonth) & "-01"
    If Not IsDate(strTemp) Then
        Call VeryMonth
        strTemp = txtEdit(Idx_txtYear) & "-" & txtEdit(Idx_txtMonth) & "-01"
        If IsDate(strTemp) = False Then
            Call VeryYear
            strTemp = txtEdit(Idx_txtYear) & "-" & txtEdit(Idx_txtMonth) & "-01"
        End If
    End If
    dtTemp1 = Format(CDate(strTemp), "yyyy-MM-01")
    dtTemp2 = DateAdd("m", 1, dtTemp1)
    GetLastDateOfMonth = Val(Format(DateAdd("d", -1, dtTemp2), "DD"))
End Function
Public Sub TxtSelAll(objTxt As Object)
    '功能：将编辑框的的文本全部选中
    '参数：objTxt=需要全选的编辑控件,该控件具有SelStart,SelLength属性
    objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    mintLastSetFocus = Index
    TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{tab}": Exit Sub
    End If
    Select Case Index
    Case Idx_txtYear  '年
        If KeyCode = vbKeyLeft Then
            If Trim(txtEdit(Index).Text) = "" Then Call VeryYear
            txtEdit(Idx_txtDay).SetFocus: Exit Sub
        End If
        If KeyCode = vbKeyRight Then
            If Trim(txtEdit(Index).Text) = "" Then Call VeryYear
            txtEdit(Idx_txtMonth).SetFocus: Exit Sub
        End If
    Case Idx_txtMonth '月
        If KeyCode = vbKeyLeft Then
            If Trim(txtEdit(Index).Text) = "" Then Call VeryMonth
            txtEdit(Idx_txtYear).SetFocus: Exit Sub
        End If
        If KeyCode = vbKeyRight Then
            If Trim(txtEdit(Index).Text) = "" Then Call VeryMonth
            txtEdit(Idx_txtDay).SetFocus: Exit Sub
        End If
        If KeyCode = vbKeyBack And txtEdit(Index).SelStart = 0 Then
            If Trim(txtEdit(Index).Text) = "" Then Call VeryMonth
            txtEdit(Idx_txtYear).SetFocus: txtEdit(Idx_txtYear).SelStart = Len(txtEdit(Idx_txtYear))
            txtEdit(Idx_txtYear).SelLength = 0
        End If
    Case Idx_txtDay '日
        If KeyCode = vbKeyLeft Then
            If Trim(txtEdit(Index).Text) = "" Then Call VeryDay
            txtEdit(Idx_txtMonth).SetFocus: Exit Sub
        End If
        If KeyCode = vbKeyRight Then
            If Trim(txtEdit(Index).Text) = "" Then Call VeryDay
            txtEdit(Idx_txtYear).SetFocus: Exit Sub
        End If
        If KeyCode = vbKeyBack And txtEdit(Index).SelStart = 0 Then
            If Trim(txtEdit(Index).Text) = "" Then Call VeryDay
            txtEdit(Idx_txtMonth).SetFocus: txtEdit(Idx_txtMonth).SelStart = Len(txtEdit(Idx_txtMonth))
            txtEdit(Idx_txtMonth).SelLength = 0
        End If
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) _
        And KeyAscii <> vbKeyBack _
        And KeyAscii <> vbKeyLeft _
        And KeyAscii <> vbKeyRight Then
        If KeyAscii = vbKeyTab Then Exit Sub
        KeyAscii = 0: Exit Sub
    End If
End Sub
Private Function VeryYear() As Boolean
    Dim blnChange As Boolean
    '验证年
    '先看是否为年为最大日期的限制
    mblnNotChange = True: blnChange = False
    If Val(txtEdit(Idx_txtYear).Text) > Val(Format(m_MaxDate, "yyyy")) Then
            txtEdit(Idx_txtYear).Text = Format(m_MaxDate, "yyyy"): blnChange = True
    ElseIf Val(txtEdit(Idx_txtYear).Text) < Val(Format(m_MinDate, "yyyy")) Then
            txtEdit(Idx_txtYear).Text = Format(m_MinDate, "yyyy"): blnChange = True
    ElseIf Trim(txtEdit(Idx_txtMonth).Text) = "" Then
            txtEdit(Idx_txtYear).Text = Format(m_Value, "YYYY"): blnChange = True
    Else '肯定正常
    End If
    If blnChange Then
        RaiseEvent Change
    End If
    mblnNotChange = False
End Function
Private Function VeryMonth() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：校证月份
    '返回：如果校证过的,则返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-06-28 14:21:22
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, blnChange As Boolean
    mblnNotChange = True: blnChange = False
    
    strTemp = txtEdit(Idx_txtYear).Text & "-" & Lpad(Trim(txtEdit(Idx_txtMonth).Text), 2, "0")
    '如果是月,先看是否为年为最大日期的限制
    If strTemp > Format(m_MaxDate, "yyyy-MM") Then
            txtEdit(Idx_txtMonth).Text = Format(m_MaxDate, "MM"): blnChange = True
    ElseIf strTemp < Format(m_MinDate, "yyyy-MM") Then
            txtEdit(Idx_txtMonth).Text = Format(m_MinDate, "MM"): blnChange = True
    ElseIf Trim(txtEdit(Idx_txtMonth).Text) = "" Then
            txtEdit(Idx_txtMonth).Text = Format(m_Value, "MM"): blnChange = True
    ElseIf Val(txtEdit(Idx_txtMonth)) > 12 Then    '输入的值比12还大,肯定以不正确,因此,强制为12
         txtEdit(Idx_txtMonth).Text = 12: blnChange = True
    ElseIf Val(txtEdit(Idx_txtMonth)) < 1 Then   '输入的值比1还小,肯定不正确,因此,强制为01
         txtEdit(Idx_txtMonth).Text = "01": blnChange = True
    Else '肯定服合要求
    End If
  
    If blnChange Then RaiseEvent Change
    
    mblnNotChange = False
End Function
Private Function VeryDay() As Boolean
    Dim dtDate As Date, strTemp As String, intMonth As Integer, blnChange As Boolean
    strTemp = txtEdit(Idx_txtYear).Text & "-" & Lpad(txtEdit(Idx_txtMonth).Text, 2, "0") & "-" & Lpad(txtEdit(Idx_txtDay).Text, 2, "0")
    mblnNotChange = True: blnChange = False
    '如果是月,先看是否为年为最大日期的限制
    If strTemp > Format(m_MaxDate, "yyyy-MM-DD") Then
            txtEdit(Idx_txtDay).Text = Format(m_MaxDate, "DD"): blnChange = True
    ElseIf strTemp < Format(m_MinDate, "yyyy-MM-DD") Then
            txtEdit(Idx_txtDay).Text = Format(m_MinDate, "DD"): blnChange = True
    ElseIf Trim(txtEdit(Idx_txtDay).Text) = "" Then
            txtEdit(Idx_txtDay).Text = Format(m_Value, "DD"): blnChange = True
    Else
        intMonth = GetLastDateOfMonth
        If intMonth < Val(txtEdit(Idx_txtDay).Text) Then '如果此月的最后一天比现在的还大,则以最后一天为准
            txtEdit(Idx_txtDay).Text = Lpad(intMonth, 2, "0"): blnChange = True
        ElseIf Val(txtEdit(Idx_txtDay).Text) <= 0 Then
            txtEdit(Idx_txtDay).Text = "01": blnChange = True
        End If
    End If
    
    If blnChange Then
        RaiseEvent Change
    End If
    VeryDay = blnChange
    mblnNotChange = False
End Function
Private Sub txtEdit_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Trim(txtEdit(Idx_txtYear).Text) = "" Then Call VeryYear
    If Trim(txtEdit(Idx_txtMonth).Text) = "" Then Call VeryMonth
    If Trim(txtEdit(Idx_txtDay).Text) = "" Then Call VeryDay
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
        Select Case Index
        Case Idx_txtYear   '年的处理
            Call VeryYear
        Case Idx_txtMonth '月的处理
            Call VeryMonth
        Case Else '日的验证
            Call VeryDay
        End Select
End Sub
  
 
Private Sub zlReSetDefaultDate()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：重新设置缺省时间
    '编制：刘兴洪
    '日期：2010-06-23 17:30:33
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    mblnNotChange = True
    txtEdit(Idx_txtYear).Text = Format(m_Value, "yyyy")
    txtEdit(Idx_txtMonth).Text = Format(m_Value, "MM")
    txtEdit(Idx_txtDay).Text = Format(m_Value, "DD")
    mblnNotChange = False
End Sub

Private Sub UserControl_GotFocus()
    Err = 0: On Error Resume Next
    txtEdit(Idx_txtYear).SetFocus
End Sub

Private Sub UserControl_Initialize()
        m_MinDate = CDate("1601-01-01")
        m_MaxDate = CDate("9999-12-31")
        m_Value = Now
End Sub
Private Sub UserControl_Paint()
    Err = 0: On Error Resume Next
     Height = picDate.Height
     With picDate
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
        'Width = .Width
    End With
    Call MoveCtrl
End Sub

Private Sub UserControl_Resize()
    Err = 0: On Error Resume Next
    ' Height = picDate.Height
     With picDate
        .Left = ScaleLeft
        .Top = ScaleTop
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With
    Call MoveCtrl
End Sub
Private Sub MoveCtrl()
    Err = 0: On Error Resume Next
    txtEdit(Idx_txtYear).Top = (picDate.ScaleHeight - txtEdit(Idx_txtYear).Height) \ 2 + 10
    txtEdit(Idx_txtMonth).Top = txtEdit(Idx_txtYear).Top
    txtEdit(Idx_txtDay).Top = txtEdit(Idx_txtYear).Top
    cmdDate.Top = picDate.ScaleTop + 5
    cmdDate.Height = IIf(picDate.ScaleHeight - cmdDate.Top < 0, 0, picDate.ScaleHeight - cmdDate.Top)
    
    lblMonthSplit.Left = txtEdit(Idx_txtYear).Left + txtEdit(Idx_txtYear).Width
    txtEdit(Idx_txtMonth).Left = lblMonthSplit.Left + lblMonthSplit.Width
    lblDaySplit.Left = txtEdit(Idx_txtMonth).Left + txtEdit(Idx_txtMonth).Width
    txtEdit(Idx_txtDay).Left = lblDaySplit.Left + lblDaySplit.Width
    lblEnd.Left = txtEdit(Idx_txtDay).Left + txtEdit(Idx_txtDay).Width
    lblMonthSplit.Top = (picDate.ScaleHeight - lblMonthSplit.Height) \ 2
    lblDaySplit.Top = lblMonthSplit.Top
    lblEnd.Top = lblMonthSplit.Top
End Sub
Private Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = Substr(strCode, 1, lngLen)
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function
Private Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取指定字串的值,字串中可以包含汉字
    '--入参数:strInfor-原串
    '         lngStart-直始位置
    '         lngLen-长度
    '--出参数:
    '--返  回:子串
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    Err = 0
    On Error GoTo ErrHand:
    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
    Exit Function
ErrHand:
    Substr = ""
End Function
 
'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=3,0,0,now
Public Property Get Value() As Date
Attribute Value.VB_Description = "当前日期"
    Dim strTemp As String
    strTemp = txtEdit(Idx_txtYear).Text & "-" & txtEdit(Idx_txtMonth).Text & "-" & txtEdit(Idx_txtDay).Text
    If IsDate(strTemp) Then '是日期
        m_Value = CDate(strTemp)
        
    Else '不是日期,可能为空的原因,只能强制修正为日期
          Call VeryYear: Call VeryMonth: Call VeryDay
          strTemp = txtEdit(Idx_txtYear).Text & "-" & txtEdit(Idx_txtMonth).Text & "-" & txtEdit(Idx_txtDay).Text
          If IsDate(strTemp) Then
                m_Value = CDate(strTemp)
          Else '如果修正不成功,则按缺省日期处理
                If m_Value > m_MaxDate Then m_Value = m_MaxDate
                If m_Value < m_MinDate Then m_Value = m_MinDate
                 Call zlReSetDefaultDate
          End If
    End If
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Date)
    m_Value = Format(New_Value, "yyyy-mm-dd")
    If m_Value > m_MaxDate Then m_Value = m_MinDate
    If m_Value < m_MinDate Then m_Value = m_MinDate
    Call zlReSetDefaultDate
    PropertyChanged "Value"
    RaiseEvent Change
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "返回/设置一个值，决定一个对象是否响应用户生成事件。"
    Enabled = UserControl.Enabled
    
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    txtEdit(Idx_txtDay).Enabled = New_Enabled
    txtEdit(Idx_txtMonth).Enabled = New_Enabled
    txtEdit(Idx_txtYear).Enabled = New_Enabled
    cmdDate.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=picDate,picDate,-1,Appearance
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "返回/设置一个对象在运行时是否以 3D 效果显示。"
    Appearance = picDate.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    picDate.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=3,0,0,9999-01-01
Public Property Get MaxDate() As Date
    MaxDate = m_MaxDate
End Property

Public Property Let MaxDate(ByVal New_MaxDate As Date)
    m_MaxDate = New_MaxDate
    If m_MaxDate < m_MinDate Then m_MinDate = m_MaxDate
    If m_Value > m_MaxDate Then m_Value = m_MaxDate
    Call zlReSetDefaultDate
    PropertyChanged "MaxDate"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=3,0,0,1601-01-01
Public Property Get MinDate() As Date
    MinDate = m_MinDate
End Property

Public Property Let MinDate(ByVal New_MinDate As Date)
    m_MinDate = New_MinDate
    If m_MaxDate < m_MinDate Then m_MaxDate = m_MinDate
    If m_Value < m_MinDate Then m_Value = m_MinDate
    Call zlReSetDefaultDate
    PropertyChanged "MinDate"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=13,0,0,YYYY-MM-DD
Public Property Get CustomFormat() As String
    CustomFormat = m_CustomFormat
End Property

Public Property Let CustomFormat(ByVal New_CustomFormat As String)
        Dim varData As Variant, varTemp As Variant
        m_CustomFormat = New_CustomFormat
        If mstrCustomFormat <> "" Then
            varData = Split(UCase(mstrCustomFormat), "MM")
            If UBound(varData) >= 1 Then
                varTemp = Split(varData(0), "YYYY")
                If UBound(varTemp) >= 1 Then lblMonthSplit = varTemp(1)
                varTemp = Split(varData(1), "DD")
                If UBound(varTemp) >= 1 Then
                    lblDaySplit = varTemp(0)
                    If varTemp(1) <> "" Then
                        lblEnd.Caption = varTemp(1)
                    End If
                    lblEnd.Visible = lblEnd.Caption <> ""
                End If
            End If
        End If
        Call UserControl_Resize
    PropertyChanged "CustomFormat"
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_MaxDate = m_def_MaxDate
    m_MinDate = m_def_MinDate
    m_CustomFormat = m_def_CustomFormat
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    txtEdit(Idx_txtDay).Enabled = UserControl.Enabled
    txtEdit(Idx_txtMonth).Enabled = UserControl.Enabled
    txtEdit(Idx_txtYear).Enabled = UserControl.Enabled
    cmdDate.Enabled = UserControl.Enabled
    picDate.Appearance = PropBag.ReadProperty("Appearance", 1)
    m_MaxDate = PropBag.ReadProperty("MaxDate", m_def_MaxDate)
    m_MinDate = PropBag.ReadProperty("MinDate", m_def_MinDate)
    m_CustomFormat = PropBag.ReadProperty("CustomFormat", m_def_CustomFormat)
    picDate.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    
    txtEdit(Idx_txtYear).BackColor = picDate.BackColor
    txtEdit(Idx_txtMonth).BackColor = picDate.BackColor
    txtEdit(Idx_txtDay).BackColor = picDate.BackColor
    lblDaySplit.BackColor = picDate.BackColor
    lblEnd.BackColor = picDate.BackColor
    lblMonthSplit.BackColor = picDate.BackColor
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Appearance", picDate.Appearance, 1)
    Call PropBag.WriteProperty("MaxDate", m_MaxDate, m_def_MaxDate)
    Call PropBag.WriteProperty("MinDate", m_MinDate, m_def_MinDate)
    Call PropBag.WriteProperty("CustomFormat", m_CustomFormat, m_def_CustomFormat)
    Call PropBag.WriteProperty("BackColor", picDate.BackColor, &H80000005)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=picDate,picDate,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = picDate.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picDate.BackColor() = New_BackColor
    txtEdit(Idx_txtYear).BackColor = New_BackColor
    txtEdit(Idx_txtMonth).BackColor = New_BackColor
    txtEdit(Idx_txtDay).BackColor = New_BackColor
    lblDaySplit.BackColor = New_BackColor
    lblEnd.BackColor = New_BackColor
    lblMonthSplit.BackColor = New_BackColor
    
    PropertyChanged "BackColor"
End Property

