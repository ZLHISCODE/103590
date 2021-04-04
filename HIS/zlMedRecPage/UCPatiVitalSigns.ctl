VERSION 5.00
Begin VB.UserControl UCPatiVitalSigns 
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   ScaleHeight     =   900
   ScaleWidth      =   6165
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   0
      Left            =   390
      TabIndex        =   1
      Top             =   0
      Width           =   555
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   1
      Left            =   1950
      TabIndex        =   2
      Top             =   0
      Width           =   555
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   2
      Left            =   4035
      TabIndex        =   3
      Top             =   0
      Width           =   555
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   3
      Left            =   405
      TabIndex        =   4
      Top             =   375
      Width           =   555
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   4
      Left            =   1980
      TabIndex        =   5
      Top             =   360
      Width           =   555
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   5
      Left            =   4050
      TabIndex        =   6
      Top             =   405
      Width           =   555
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   6
      Left            =   4770
      TabIndex        =   7
      Top             =   405
      Width           =   555
   End
   Begin VB.Frame fraCboBorder 
      Height          =   300
      Left            =   5295
      TabIndex        =   0
      Top             =   405
      Width           =   765
      Begin VB.ComboBox cboBpUnit 
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
         ItemData        =   "UCPatiVitalSigns.ctx":0000
         Left            =   -120
         List            =   "UCPatiVitalSigns.ctx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Label lblName 
      Caption         =   "身高"
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   15
      Width           =   405
   End
   Begin VB.Label lblName 
      Caption         =   "体重"
      Height          =   180
      Index           =   1
      Left            =   1560
      TabIndex        =   19
      Top             =   60
      Width           =   400
   End
   Begin VB.Label lblName 
      Caption         =   "体温"
      Height          =   180
      Index           =   2
      Left            =   3615
      TabIndex        =   18
      Top             =   60
      Width           =   400
   End
   Begin VB.Label lblName 
      Caption         =   "脉搏"
      Height          =   180
      Index           =   3
      Left            =   30
      TabIndex        =   17
      Top             =   450
      Width           =   405
   End
   Begin VB.Label lblName 
      Caption         =   "呼吸"
      Height          =   180
      Index           =   4
      Left            =   1620
      TabIndex        =   16
      Top             =   405
      Width           =   400
   End
   Begin VB.Label lblName 
      Caption         =   "血压"
      Height          =   180
      Index           =   5
      Left            =   3570
      TabIndex        =   15
      Top             =   390
      Width           =   405
   End
   Begin VB.Label lblName 
      Caption         =   "/"
      Height          =   180
      Index           =   6
      Left            =   4650
      TabIndex        =   14
      Top             =   450
      Width           =   80
   End
   Begin VB.Label lblUnit 
      Caption         =   "cm"
      Height          =   180
      Index           =   0
      Left            =   1095
      TabIndex        =   13
      Top             =   60
      Width           =   450
   End
   Begin VB.Label lblUnit 
      Caption         =   "kg"
      Height          =   180
      Index           =   1
      Left            =   2730
      TabIndex        =   12
      Top             =   75
      Width           =   450
   End
   Begin VB.Label lblUnit 
      Caption         =   "℃"
      Height          =   180
      Index           =   2
      Left            =   4890
      TabIndex        =   11
      Top             =   60
      Width           =   450
   End
   Begin VB.Label lblUnit 
      Caption         =   "次/分"
      Height          =   180
      Index           =   3
      Left            =   1095
      TabIndex        =   10
      Top             =   420
      Width           =   450
   End
   Begin VB.Label lblUnit 
      Caption         =   "次/分"
      Height          =   180
      Index           =   4
      Left            =   2790
      TabIndex        =   9
      Top             =   420
      Width           =   450
   End
End
Attribute VB_Name = "UCPatiVitalSigns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Change(ByVal int序号 As Integer)
'    txt身高 = 0
'    txt体重 = 1
'    txt体温 = 2
'    txt脉搏 = 3
'    txt呼吸 = 4
'    txt收缩压 = 5
'    txt舒张压 = 6
'    血压单位 = 7

Private Enum enum_txtInfo '文本框
    txt身高 = 0
    txt体重 = 1
    txt体温 = 2
    txt脉搏 = 3
    txt呼吸 = 4
    txt收缩压 = 5
    txt舒张压 = 6
End Enum

Private Enum enum_lblName '文本框名称--项目名称
    lblN身高 = 0
    lblN体重 = 1
    lblN体温 = 2
    lblN脉搏 = 3
    lblN呼吸 = 4
    lblN血压 = 5
    lblN血压分割线 = 6
End Enum

Private Enum enum_lblUnit  '文本框单位--项目单位
    lblU身高 = 0
    lblU体重 = 1
    lblU体温 = 2
    lblU脉搏 = 3
    lblU呼吸 = 4
End Enum
 
Public Enum enum_Style '文本框的风格 默认0-TextBox
    TextBox = 0
    Underline = 1
End Enum
Private mEnumStyle As enum_Style

Public Enum enum_ShowMode '行数，目前只提供1行和2行形式
    OneRow = 0
    TwoRow = 1
End Enum
Private mEnumShowMode As enum_ShowMode

Private mXDis As Long '水平方向各项目之间的间隔
Private mYDis As Long '竖直方向各项目之间的间隔，仅在两行显示模式下生效
Private mLabToTxt As Long '标签名和文本框的距离
Private mcolForeColor As OLE_COLOR
Private mstrTag As String
Private mcol范围 As Collection '各项的取值范围
Private mstrValues As String

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    Dim i As Long
    
    UserControl.Enabled = NewValue
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).Enabled = NewValue
        txtInfo(i).BackColor = IIf(NewValue, vbWindowBackground, vbButtonFace)
    Next
    cboBpUnit.Enabled = NewValue
    cboBpUnit.BackColor = IIf(NewValue, vbWindowBackground, vbButtonFace)
    
    PropertyChanged "Enabled"
End Property

Public Property Get ControlLock() As Boolean
    ControlLock = txtInfo(txt身高).Locked
End Property

Public Property Let ControlLock(ByVal NewValue As Boolean)
    Dim i As Long
    
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).Locked = NewValue
        txtInfo(i).TabStop = Not NewValue
        If txtInfo(i).Enabled Then txtInfo(i).BackColor = IIf(NewValue, vbButtonFace, vbWindowBackground)
    Next
    cboBpUnit.Locked = NewValue
    
    PropertyChanged "ControlLock"
End Property

Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Dim i As Long
    
    Set UserControl.Font = New_Font
    
    For i = 0 To txtInfo.Count - 1
        Set txtInfo(i).Font = New_Font
    Next
    
    For i = 0 To lblName.Count - 1
        Set lblName(i).Font = New_Font
    Next
    
    For i = 0 To lblUnit.Count - 1
        Set lblUnit(i).Font = New_Font
    Next
    
    Set cboBpUnit.Font = New_Font
    
    Call UserControl_Resize
    
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mcolForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Dim i As Long
    
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).ForeColor() = New_ForeColor
    Next
    mcolForeColor = New_ForeColor
    
    PropertyChanged "ForeColor"
End Property

Public Property Get TextBackColor() As OLE_COLOR
    TextBackColor = txtInfo(txt身高).BackColor
End Property

Public Property Let TextBackColor(ByVal New_BackColor As OLE_COLOR)
    Dim i As Long
    
    For i = 0 To txtInfo.Count - 1
       txtInfo(i).BackColor = New_BackColor
    Next
    
    cboBpUnit.BackColor = New_BackColor
    
    PropertyChanged "TextBackColor"
End Property

Public Property Get LblBackColor() As OLE_COLOR
    LblBackColor = lblName(lblN身高).BackColor
End Property

Public Property Let LblBackColor(ByVal New_BackColor As OLE_COLOR)
    Dim i As Long

    For i = 0 To lblName.Count - 1
        lblName(i).BackColor = New_BackColor
    Next
    
    For i = 0 To lblUnit.Count - 1
        lblUnit(i).BackColor = New_BackColor
    Next
    
    PropertyChanged "TextBackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
'    Call SetLine(Me.Style)
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    Call SetLine(Me.Style)
    PropertyChanged "BackColor"
End Property

Public Property Get MaxLength() As Long
    MaxLength = txtInfo(txt身高).MaxLength
End Property

Public Property Let MaxLength(ByVal vNewValue As Long)
    Dim i As Long
    
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).MaxLength = vNewValue
    Next
    
    PropertyChanged "MaxLength"
End Property

Public Property Get Tag() As String
    Tag = mstrTag
End Property

Public Property Let Tag(ByVal vNewValue As String)
    mstrTag = vNewValue
    PropertyChanged "Tag"
End Property

Public Property Get ShowMode() As enum_ShowMode
    ShowMode = mEnumShowMode
End Property

Public Property Let ShowMode(ByVal vNewValue As enum_ShowMode)
    Dim i As Long
    
    mEnumShowMode = vNewValue
    Call UserControl_Resize
    
    PropertyChanged "ShowMode"
End Property


Public Property Get XDis() As Long
    XDis = mXDis
End Property

Public Property Let XDis(ByVal vNewValue As Long)
    Dim i As Long
    
    mXDis = vNewValue

    Call UserControl_Resize
    Call SetLine(Me.Style)
    PropertyChanged "XDis"
End Property

Public Property Get YDis() As Long
    YDis = mYDis
End Property

Public Property Let YDis(ByVal vNewValue As Long)
    Dim i As Long
    
    mYDis = vNewValue

    Call UserControl_Resize
    Call SetLine(Me.Style)
    PropertyChanged "YDis"
End Property

Public Property Get LabToTxt() As Long
    LabToTxt = mLabToTxt
End Property

Public Property Let LabToTxt(ByVal vNewValue As Long)
    Dim i As Long
    
    mLabToTxt = vNewValue

    Call UserControl_Resize
    Call SetLine(Me.Style)
    PropertyChanged "LabToTxt"
End Property

Public Property Get Style() As enum_Style
    Style = mEnumStyle
End Property

Public Property Let Style(ByVal vNewValue As enum_Style)
    Dim i As Long
    
    mEnumStyle = vNewValue
    
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).BorderStyle = IIf(vNewValue = 0, 1, 0)
    Next
    Call UserControl_Resize
    Call SetLine(vNewValue)
    PropertyChanged "Style"
End Property

Public Property Get value身高() As String
    value身高 = txtInfo(txt身高).Text
End Property

Public Property Let value身高(ByVal vNewValue As String)
    txtInfo(txt身高).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value身高"
End Property

Public Property Get value体重() As String
    value体重 = txtInfo(txt体重).Text
End Property

Public Property Let value体重(ByVal vNewValue As String)
    txtInfo(txt体重).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value体重"
End Property

Public Property Get value体温() As String
    value体温 = txtInfo(txt体温).Text
End Property

Public Property Let value体温(ByVal vNewValue As String)
    txtInfo(txt体温).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value体温"
End Property

Public Property Get value脉搏() As String
    value脉搏 = txtInfo(txt脉搏).Text
End Property

Public Property Let value脉搏(ByVal vNewValue As String)
    txtInfo(txt脉搏).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value脉搏"
End Property

Public Property Get value呼吸() As String
    value呼吸 = txtInfo(txt呼吸).Text
End Property

Public Property Let value呼吸(ByVal vNewValue As String)
    txtInfo(txt呼吸).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value呼吸"
End Property

Public Property Get value收缩压() As String
    value收缩压 = txtInfo(txt收缩压).Text
End Property

Public Property Let value收缩压(ByVal vNewValue As String)
    txtInfo(txt收缩压).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value收缩压"
End Property

Public Property Get value舒张压() As String
    value舒张压 = txtInfo(txt舒张压).Text
End Property

Public Property Let value舒张压(ByVal vNewValue As String)
    txtInfo(txt舒张压).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value舒张压"
End Property

Public Property Get value血压单位() As String
    value血压单位 = cboBpUnit.Text
End Property

Public Property Let value血压单位(ByVal vNewValue As String)
    Call zlControl.CboLocate(cboBpUnit, vNewValue)
    PropertyChanged "value血压单位"
End Property

Private Sub SetLine(ByVal lngStyle As Long)
'功能：设置下划线
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    UserControl.Cls
    For i = 0 To txtInfo.Count - 1
        If lngStyle = Underline Then
            x1 = txtInfo(i).Left
            y1 = txtInfo(i).Top + txtInfo(i).Height
            x2 = txtInfo(i).Left + txtInfo(i).Width - 60
            y2 = y1
            UserControl.Line (x1, y1)-(x2, y2), &H808080
        End If
    Next
    If lngStyle = Underline Then
        x1 = fraCboBorder.Left
        y1 = fraCboBorder.Top + fraCboBorder.Height
        x2 = fraCboBorder.Left + fraCboBorder.Width
        y2 = y1
        UserControl.Line (x1, y1)-(x2, y2), &H808080
    End If
End Sub

Private Sub cboBpUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strMask As String
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        
        Select Case Index
            Case txt身高, txt呼吸, txt脉搏
                strMask = "1234567890"
            Case txt收缩压, txt舒张压
                If cboBpUnit.Text = "mmHg" Then
                    strMask = "1234567890"
                Else
                    strMask = "1234567890."
                End If
            Case txt体重, txt体温
                strMask = "1234567890."
        End Select
        
        If InStr(strMask, Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
        
    End If
End Sub

Private Sub UserControl_Paint()
    Call SetLine(Me.Style)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Me.Style = PropBag.ReadProperty("Style", enum_Style.TextBox)
    Me.ShowMode = PropBag.ReadProperty("ShowMode", enum_ShowMode.TwoRow)
    Me.ControlLock = PropBag.ReadProperty("ControlLock", False)
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    Me.TextBackColor = PropBag.ReadProperty("TextBackColor", &H80000005)
    Me.LblBackColor = PropBag.ReadProperty("LblBackColor", &H8000000F)
    Set Me.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Me.ForeColor = PropBag.ReadProperty("ForeColor", &H80000000)
    Me.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Me.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Me.XDis = PropBag.ReadProperty("XDis", 20)
    Me.YDis = PropBag.ReadProperty("YDis", 10)
    Me.LabToTxt = PropBag.ReadProperty("LabToTxt", 10)
    
    Me.value身高 = PropBag.ReadProperty("value身高", "")
    Me.value体重 = PropBag.ReadProperty("value体重", "")
    Me.value体温 = PropBag.ReadProperty("value体温", "")
    Me.value脉搏 = PropBag.ReadProperty("value脉搏", "")
    Me.value呼吸 = PropBag.ReadProperty("value呼吸", "")
    Me.value收缩压 = PropBag.ReadProperty("value收缩压", "")
    Me.value舒张压 = PropBag.ReadProperty("value舒张压", "")
    Me.value血压单位 = PropBag.ReadProperty("value血压单位", "")
    Me.Tag = PropBag.ReadProperty("Tag", "")
End Sub

Private Sub UserControl_Resize()
'功能: 设置控件大小位置
    On Error Resume Next
    Dim lngHeight As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim i As Integer
    
    '取字体自适应长宽
    lngHeight = UserControl.TextHeight("中")
    
    lngTop = 0
    lngLeft = 0
    For i = 0 To 5
        If i = 3 And mEnumShowMode = TwoRow Then
            lngTop = txtInfo(0).Height + Me.YDis
            lngLeft = 0
        End If
        lblName(i).Move lngLeft, lngTop, UserControl.TextWidth(IIf(Me.Style = Underline, "中联:", "中联")), lngHeight
        lngLeft = lngLeft + lblName(i).Width + IIf(Me.Style = TextBox, Me.LabToTxt, 0)
        txtInfo(i).Move lngLeft, lngTop, UserControl.TextWidth("中联信"), lngHeight
        lngLeft = lngLeft + txtInfo(i).Width + IIf(Me.Style = TextBox, Me.LabToTxt, 0)
        lblUnit(i).Move lngLeft, lngTop, UserControl.TextWidth("中/联"), lngHeight
        lngLeft = lngLeft + lblUnit(i).Width + Me.XDis
    Next
    
    If Me.Style = Underline Then
        For i = 0 To 5
            If InStr(lblName(i).Caption, ":") = 0 Then
                lblName(i).Caption = lblName(i).Caption & ":"
            End If
        Next
    End If
    
    lblName(lblN血压分割线).Top = lngTop
    lblName(lblN血压分割线).Width = UserControl.TextWidth("/")
    lblName(lblN血压分割线).Height = lngHeight
    lblName(lblN血压分割线).Left = txtInfo(txt收缩压).Left + txtInfo(txt收缩压).Width
    
    txtInfo(txt舒张压).Move lblName(lblN血压分割线).Left + lblName(lblN血压分割线).Width, lngTop, UserControl.TextWidth("中联信"), lngHeight
    
    If Me.Style = TextBox Then
        txtInfo(txt舒张压).Height = IIf(txtInfo(txt舒张压).Height < 300, 300, txtInfo(txt舒张压).Height)
        For i = 0 To 5
            txtInfo(i).Height = IIf(txtInfo(i).Height < 300, 300, txtInfo(i).Height)
            lblName(i).Top = (txtInfo(i).Height - lblName(i).Height) / 2 + txtInfo(i).Top
            lblUnit(i).Top = (txtInfo(i).Height - lblUnit(i).Height) / 2 + txtInfo(i).Top
        Next
        lblName(lblN血压分割线).Top = (txtInfo(txt舒张压).Height - lblName(lblN血压分割线).Height) / 2 + txtInfo(txt舒张压).Top
    End If
    
    cboBpUnit.Left = txtInfo(txt舒张压).Left + txtInfo(txt舒张压).Width
    cboBpUnit.Top = txtInfo(txt舒张压).Top
    cboBpUnit.Width = UserControl.TextWidth("mmHg") + 400
    cboBpUnit.Height = lngHeight * 2
    
    cboBpUnit.Left = IIf(Me.Style = Underline, -30, 0)
    cboBpUnit.Top = IIf(Me.Style = Underline, -30, 0)
    
    If Me.Style = Underline Then
        fraCboBorder.Height = IIf(cboBpUnit.Height <= 300, 240, cboBpUnit.Height - 40)
        fraCboBorder.Top = txtInfo(txt舒张压).Top
        fraCboBorder.Width = UserControl.TextWidth("mmHg") + 350
    Else
        fraCboBorder.Height = txtInfo(0).Height
        fraCboBorder.Top = txtInfo(txt舒张压).Top
        fraCboBorder.Width = cboBpUnit.Width
    End If
    
    fraCboBorder.Left = txtInfo(txt舒张压).Left + txtInfo(txt舒张压).Width
    
    UserControl.Width = fraCboBorder.Left + fraCboBorder.Width
    UserControl.Height = txtInfo(txt舒张压).Top + txtInfo(txt舒张压).Height + 100
End Sub

Private Sub UserControl_Show()
    Call SetLine(Me.Style)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ControlLock", txtInfo(txt身高).Locked, False)
    Call PropBag.WriteProperty("Enabled", Me.Enabled, True)
    Call PropBag.WriteProperty("TextBackColor", txtInfo(txt身高).BackColor, &H80000000)
    Call PropBag.WriteProperty("LblBackColor", lblName(lblN身高).BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", Me.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", mcolForeColor, &H80000000)
    Call PropBag.WriteProperty("BackColor", Me.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ShowMode", Me.ShowMode, enum_ShowMode.TwoRow)
    Call PropBag.WriteProperty("Style", Me.Style, enum_Style.TextBox)
    Call PropBag.WriteProperty("MaxLength", txtInfo(txt身高).MaxLength, 0)
    Call PropBag.WriteProperty("XDis", mXDis, 20)
    Call PropBag.WriteProperty("YDis", mYDis, 10)
    Call PropBag.WriteProperty("LabToTxt", mLabToTxt, 10)
    
    Call PropBag.WriteProperty("value身高", Me.value身高, "")
    Call PropBag.WriteProperty("value体重", Me.value体重, "")
    Call PropBag.WriteProperty("value呼吸", Me.value呼吸, "")
    Call PropBag.WriteProperty("value脉搏", Me.value脉搏, "")
    Call PropBag.WriteProperty("value体温", Me.value体温, "")
    Call PropBag.WriteProperty("value收缩压", Me.value收缩压, "")
    Call PropBag.WriteProperty("value舒张压", Me.value舒张压, "")
    Call PropBag.WriteProperty("value血压单位", Me.value血压单位, "")
    Call PropBag.WriteProperty("Tag", Me.Tag, "")
End Sub

Public Function LoadPatiVitalSigns(ByVal lng病人ID As Long, ByVal lng挂号id As Long)
'功能：加载记录内容到对应的文本框中，主要将数据保存于文本框的Tag值中
'注意：体重和血压等的输入时暂不严格控制，但可以保证结果是在规定的范围内
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Integer
    Dim blnSave As Boolean
    
    Call ClearData
    
    With cboBpUnit
        .Clear
        .AddItem "mmHg"
        .AddItem "Kpa"
    End With
    
    Set mcol范围 = New Collection
    
    strSQL = "Select ID, 中文名, 长度, 单位, 数值域, 小数 From 诊治所见项目" & _
        " Where 分类id = 7 And 中文名 In ('体温', '脉搏', '收缩压', '舒张压', '体重', '身高', '呼吸')"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatVitalSigns")
    
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
        
            strTmp = "" & rsTmp!数值域
            If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
            mcol范围.Add strTmp, "_" & rsTmp!ID
            strTmp = Replace(strTmp, ";", " - ")
            If Val("" & rsTmp!小数) = 0 Then
                strTmp = "范围为 " & strTmp & " 的数。"
            Else
                strTmp = "范围为 " & strTmp & " 之间的数，可含" & rsTmp!小数 & "位小数，最多可输入" & rsTmp!长度 & "个字符。"
            End If
            Select Case rsTmp!中文名
                Case "身高"
                    lblName(lblN身高).Tag = rsTmp!ID
                    lblUnit(lblU身高).Caption = rsTmp!单位
                    txtInfo(txt身高).MaxLength = rsTmp!长度
                    txtInfo(txt身高).ToolTipText = "身高" & strTmp
                Case "体重"
                    lblName(lblN体重).Tag = rsTmp!ID
                    lblUnit(lblU体重).Caption = rsTmp!单位
                    txtInfo(txt体重).MaxLength = 5 'rsTmp!长度
                    txtInfo(txt体重).ToolTipText = "体重" & strTmp
                Case "体温"
                    lblName(lblN体温).Tag = rsTmp!ID
                    lblUnit(lblU体温).Caption = rsTmp!单位
                    txtInfo(txt体温).MaxLength = rsTmp!长度
                    txtInfo(txt体温).ToolTipText = "体温" & strTmp
                Case "呼吸"
                    lblName(lblN呼吸).Tag = rsTmp!ID
                    lblUnit(lblU呼吸).Caption = rsTmp!单位
                    txtInfo(txt呼吸).MaxLength = rsTmp!长度
                    txtInfo(txt呼吸).ToolTipText = "呼吸" & strTmp
                Case "脉搏"
                    lblName(lblN脉搏).Tag = rsTmp!ID
                    lblUnit(lblU脉搏).Caption = rsTmp!单位
                    txtInfo(txt脉搏).MaxLength = rsTmp!长度
                    txtInfo(txt脉搏).ToolTipText = "脉搏" & strTmp
                Case "收缩压"
                    lblName(lblN血压).Tag = rsTmp!ID
                    Call zlControl.CboSetIndex(cboBpUnit.hwnd, IIf("" & rsTmp!单位 = "mmHg", 0, 1))
                    txtInfo(txt收缩压).MaxLength = 5 'rsTmp!长度
                    txtInfo(txt收缩压).ToolTipText = "收缩压" & strTmp
                Case "舒张压"
                    lblName(lblN血压分割线).Tag = rsTmp!ID
                    Call zlControl.CboSetIndex(cboBpUnit.hwnd, IIf("" & rsTmp!单位 = "mmHg", 0, 1))
                    txtInfo(txt舒张压).MaxLength = 5 'rsTmp!长度
                    txtInfo(txt舒张压).ToolTipText = "舒张压" & strTmp
            End Select
            rsTmp.MoveNext
        Next
    End If
    cboBpUnit.Tag = cboBpUnit.List(cboBpUnit.ListIndex)
    
    strSQL = "Select b.项目单位, b.项目名称, b.记录内容" & _
        " From 病人护理记录 A, 病人护理内容 B Where a.Id = b.记录id And a.病人id = [1] And a.主页id = [2]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatVitalSigns", lng病人ID, lng挂号id)
        
    If rsTmp.RecordCount <= 0 Then
        strSQL = "Select null as 项目单位, 信息名 As 项目名称, 信息值 As 记录内容 From 病人信息从表 Where 病人id = [1] And 就诊id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatVitalSigns", lng病人ID, lng挂号id)
        If rsTmp.RecordCount > 0 Then blnSave = True
    End If
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            Select Case rsTmp!项目名称
                Case "身高"
                    txtInfo(txt身高).Text = rsTmp!记录内容
                    txtInfo(txt身高).Tag = rsTmp!记录内容
                Case "体重"
                    txtInfo(txt体重).Text = rsTmp!记录内容
                    txtInfo(txt体重).Tag = rsTmp!记录内容
                Case "体温"
                    txtInfo(txt体温).Text = rsTmp!记录内容
                    txtInfo(txt体温).Tag = rsTmp!记录内容
                Case "呼吸"
                    txtInfo(txt呼吸).Text = rsTmp!记录内容
                    txtInfo(txt呼吸).Tag = rsTmp!记录内容
                Case "脉搏"
                    txtInfo(txt脉搏).Text = rsTmp!记录内容
                    txtInfo(txt脉搏).Tag = rsTmp!记录内容
                Case "收缩压"
                    txtInfo(txt收缩压).Text = rsTmp!记录内容
                    txtInfo(txt收缩压).Tag = rsTmp!记录内容
                    If Not IsNull(rsTmp!项目单位) Then Call zlControl.CboSetIndex(cboBpUnit.hwnd, IIf("" & rsTmp!项目单位 = "mmHg", 0, 1))
                Case "舒张压"
                    txtInfo(txt舒张压).Text = rsTmp!记录内容
                    txtInfo(txt舒张压).Tag = rsTmp!记录内容
                    If Not IsNull(rsTmp!项目单位) Then Call zlControl.CboSetIndex(cboBpUnit.hwnd, IIf("" & rsTmp!项目单位 = "mmHg", 0, 1))
                Case "血压单位"
                    If Not IsNull(rsTmp!记录内容) Then Call zlControl.CboSetIndex(cboBpUnit.hwnd, IIf("" & rsTmp!记录内容 = "mmHg", 0, 1))
            End Select
            rsTmp.MoveNext
        Next
    End If
    mstrValues = ""
    If Not blnSave Then
        For i = 0 To 4
            mstrValues = IIf(mstrValues <> "", mstrValues, "") & lblName(i).Tag & "|" & FormatNum(Val(txtInfo(i).Text), 2) & "|" & lblUnit(i).Caption & ","
        Next
        mstrValues = mstrValues & lblName(5).Tag & "|" & FormatNum(Val(txtInfo(5).Text), 2) & "|" & cboBpUnit.Text & "," & _
            lblName(6).Tag & "|" & FormatNum(Val(txtInfo(6).Text), 2) & "|" & cboBpUnit.Text
    End If
    If cboBpUnit.ListIndex = -1 Then Call zlControl.CboSetIndex(cboBpUnit, 0)
    Call BpRange(cboBpUnit.List(cboBpUnit.ListIndex))
    cboBpUnit.Tag = cboBpUnit.List(cboBpUnit.ListIndex)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GiveUpSave()
'功能：放弃修改值不保存，则恢复原值
    Dim i As Integer
    For i = 0 To 6
        txtInfo(i).Text = txtInfo(i).Tag
    Next
End Function
Public Function GetSaveSQL(ByVal lng病人ID As Long, ByVal lng挂号id As Long) As String
'功能：返回门诊生命体征填写的SQL
    Dim strTmp As String
    Dim strSQL As String
    Dim i As Integer
    For i = 0 To 4
        strTmp = IIf(strTmp <> "", strTmp, "") & lblName(i).Tag & "|" & FormatNum(Val(txtInfo(i).Text), 2) & "|" & lblUnit(i).Caption & ","
    Next
    strTmp = strTmp & lblName(5).Tag & "|" & FormatNum(Val(txtInfo(5).Text), 2) & "|" & cboBpUnit.Text & "," & _
        lblName(6).Tag & "|" & FormatNum(Val(txtInfo(6).Text), 2) & "|" & cboBpUnit.Text
    If strTmp <> mstrValues Then
        GetSaveSQL = "Zl_门诊生命体征_Update(" & lng病人ID & "," & lng挂号id & ",'" & strTmp & "')"
    End If
End Function

Public Function ClearData()
'功能：切换病人时清空数据，清空所有文本框的值和Tag值
    Dim i As Integer
    
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).Text = "" '输入的值
        txtInfo(i).Tag = "" '原值
    Next
    
    For i = 0 To lblName.Count - 1
        lblName(i).Tag = "" ' 项目id
    Next
    
End Function

Private Sub BpRange(ByVal str单位 As String)
'功能：血压单位变化后取值范围随之变化 '收缩压--5 '舒张压--6
    Dim strTmp As String
    Dim dblMin As Double
    Dim dblMax As Double
    Dim i As Integer
    
    If cboBpUnit.Tag <> str单位 Then
        If str单位 = "mmHg" Then
            For i = 5 To 6
                strTmp = mcol范围("_" & lblName(i).Tag)
                If InStr(strTmp, ";") > 0 Then
                    dblMin = Val(Split(strTmp, ";")(0))
                    dblMax = Val(Split(strTmp, ";")(1))
                    
                    dblMin = Round(dblMin * 10 * 3 / 4)
                    dblMax = Round(dblMax * 10 * 3 / 4)
                    txtInfo(i).ToolTipText = IIf(i = 5, "收缩", "舒张") & "压范围为 " & dblMin & " - " & dblMax & str单位
                    mcol范围.Remove ("_" & lblName(i).Tag)
                    mcol范围.Add dblMin & ";" & dblMax, "_" & lblName(i).Tag
                End If
            Next
        Else
            For i = 5 To 6
                strTmp = mcol范围("_" & lblName(i).Tag)
                If InStr(strTmp, ";") > 0 Then
                    dblMin = Val(Split(strTmp, ";")(0))
                    dblMax = Val(Split(strTmp, ";")(1))
                    
                    dblMin = Format(dblMin * 4 / 3 / 10, "#0.00")
                    dblMax = Format(dblMax * 4 / 3 / 10, "#0.00")
                    txtInfo(i).ToolTipText = IIf(i = 5, "收缩", "舒张") & "压范围为 " & dblMin & " - " & dblMax & str单位
                    mcol范围.Remove ("_" & lblName(i).Tag)
                    mcol范围.Add dblMin & ";" & dblMax, "_" & lblName(i).Tag
                End If
            Next
        End If
    End If
End Sub

Private Sub cboBpUnit_Click()
'血压单位换算
    If cboBpUnit.List(cboBpUnit.ListIndex) <> cboBpUnit.Tag Then
        If cboBpUnit.List(cboBpUnit.ListIndex) = "mmHg" Then
            'Kpa转换到mmHg 乘10再乘3减半再减半(mmHg四舍五入)
            If txtInfo(txt舒张压).Text <> "" Then
                txtInfo(txt舒张压).Text = Round(Val(txtInfo(txt舒张压).Text) * 10 * 3 / 4)
            End If
            If txtInfo(txt收缩压).Text <> "" Then
                txtInfo(txt收缩压).Text = Round(Val(txtInfo(txt收缩压).Text) * 10 * 3 / 4)
            End If
        Else
            'mmHg转换到Kpa 加倍加倍除3再除10(Kpa保留两位小数)
            If txtInfo(txt舒张压).Text <> "" Then
                txtInfo(txt舒张压).Text = Format(Val(txtInfo(txt舒张压).Text) * 4 / 3 / 10, "#0.00")
            End If
            If txtInfo(txt收缩压).Text <> "" Then
                txtInfo(txt收缩压).Text = Format(Val(txtInfo(txt收缩压).Text) * 4 / 3 / 10, "#0.00")
            End If
        End If
        Call BpRange(cboBpUnit.List(cboBpUnit.ListIndex))
        cboBpUnit.Tag = cboBpUnit.List(cboBpUnit.ListIndex)
        RaiseEvent Change(7)
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If txtInfo(Index).Text = txtInfo(Index).Tag Then Exit Sub
    RaiseEvent Change(Index)
End Sub

Private Sub cboBpUnit_Change()
    RaiseEvent Change(7)
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
'判断范围值
    Dim strTmp As String
    Dim dblMin As Double
    Dim dblMax As Double
    
    If txtInfo(Index).Text <> "" Then
        If Not IsNumeric(txtInfo(Index).Text) Then
            MsgBox "输入内容必须是数字，" & txtInfo(Index).ToolTipText, vbInformation, "中联软件"
            txtInfo(Index).Text = txtInfo(Index).Tag
            Cancel = True
            Call zlControl.TxtSelAll(txtInfo(Index))
            Exit Sub
        End If
        
        strTmp = mcol范围("_" & lblName(Index).Tag)
        If InStr(strTmp, ";") > 0 Then
            dblMin = Val(Split(strTmp, ";")(0))
            dblMax = Val(Split(strTmp, ";")(1))
            
            If Val(txtInfo(Index).Text) > dblMax Or Val(txtInfo(Index).Text) < dblMin Then
                If MsgBox("输入内容未在指定范围内，" & txtInfo(Index).ToolTipText & "是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, "中联软件") = vbNo Then
                    txtInfo(Index).Text = txtInfo(Index).Tag
                    Cancel = True
                    Call zlControl.TxtSelAll(txtInfo(Index))
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

Private Function FormatNum(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatNum = strNumber
End Function
