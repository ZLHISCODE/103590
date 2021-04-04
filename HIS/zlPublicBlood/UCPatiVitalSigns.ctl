VERSION 5.00
Begin VB.UserControl UCPatiVitalSigns 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6930
   ScaleHeight     =   375
   ScaleWidth      =   6930
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   0
      Left            =   450
      TabIndex        =   1
      Top             =   0
      Width           =   555
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   1
      Left            =   1725
      TabIndex        =   2
      Top             =   0
      Width           =   555
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   2
      Left            =   3360
      TabIndex        =   3
      Top             =   0
      Width           =   555
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   3
      Left            =   4995
      TabIndex        =   4
      Top             =   15
      Width           =   555
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   4
      Left            =   5715
      TabIndex        =   5
      Top             =   0
      Width           =   555
   End
   Begin VB.Frame fraCboBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6240
      TabIndex        =   0
      Top             =   0
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
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "体温"
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   14
      Top             =   60
      Width           =   360
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "脉搏"
      Height          =   180
      Index           =   1
      Left            =   1350
      TabIndex        =   13
      Top             =   60
      Width           =   360
   End
   Begin VB.Label lblName 
      Caption         =   "呼吸"
      Height          =   180
      Index           =   2
      Left            =   2970
      TabIndex        =   12
      Top             =   60
      Width           =   405
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "血压"
      Height          =   180
      Index           =   3
      Left            =   4515
      TabIndex        =   11
      Top             =   60
      Width           =   360
   End
   Begin VB.Label lblName 
      Caption         =   "/"
      Height          =   180
      Index           =   4
      Left            =   5595
      TabIndex        =   10
      Top             =   60
      Width           =   75
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "℃"
      Height          =   180
      Index           =   0
      Left            =   1050
      TabIndex        =   9
      Top             =   45
      Width           =   180
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "次/分"
      Height          =   180
      Index           =   1
      Left            =   2385
      TabIndex        =   8
      Top             =   45
      Width           =   450
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "次/分"
      Height          =   180
      Index           =   2
      Left            =   4020
      TabIndex        =   7
      Top             =   45
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

Private Enum E_ITEM_INDEX
    I体温 = 0
    I脉搏 = 1
    I呼吸 = 2
    I收缩压 = 3
    I舒张压 = 4
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
Private mblnColon As Boolean '标签后面是否有冒号
Private mlng收发ID As Long
Private mint性质 As Integer
Private mstrPreState As String '前一状态
Private mblnNoCheck As Boolean '不进实时保存判断
Private mblnSaveNow As Boolean '是否立即保存

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
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
    ControlLock = txtInfo(I体温).locked
End Property

Public Property Let ControlLock(ByVal NewValue As Boolean)
    Dim i As Long
    
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).locked = NewValue
        txtInfo(i).TabStop = Not NewValue
        If txtInfo(i).Enabled Then txtInfo(i).BackColor = IIf(NewValue, vbButtonFace, vbWindowBackground)
        If NewValue = True Then txtInfo(i).SelStart = 0: txtInfo(i).SelLength = 0
    Next
    cboBpUnit.locked = NewValue
    cboBpUnit.TabStop = Not NewValue
    If cboBpUnit.locked Then cboBpUnit.BackColor = IIf(NewValue, vbButtonFace, vbWindowBackground)
    
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
    Call SetLine(Me.Style)
    
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
    TextBackColor = txtInfo(I体温).BackColor
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
    LblBackColor = lblName(I体温).BackColor
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
    MaxLength = txtInfo(I体温).MaxLength
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


Public Property Get HaveColon() As Boolean
    HaveColon = mblnColon
End Property

Public Property Let HaveColon(ByVal vNewValue As Boolean)
    Dim i As Long
    
    mblnColon = vNewValue

    Call UserControl_Resize
    Call SetLine(Me.Style)
    PropertyChanged "HaveColon"
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

Public Property Get value体温() As String
    value体温 = txtInfo(I体温).Text
End Property

Public Property Let value体温(ByVal vNewValue As String)
    txtInfo(I体温).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value体温"
End Property

Public Property Get value脉搏() As String
    value脉搏 = txtInfo(I脉搏).Text
End Property

Public Property Let value脉搏(ByVal vNewValue As String)
    txtInfo(I脉搏).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value脉搏"
End Property

Public Property Get value呼吸() As String
    value呼吸 = txtInfo(I呼吸).Text
End Property

Public Property Let value呼吸(ByVal vNewValue As String)
    txtInfo(I呼吸).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value呼吸"
End Property

Public Property Get value收缩压() As String
    value收缩压 = txtInfo(I收缩压).Text
End Property

Public Property Let value收缩压(ByVal vNewValue As String)
    txtInfo(I收缩压).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value收缩压"
End Property

Public Property Get value舒张压() As String
    value舒张压 = txtInfo(I舒张压).Text
End Property

Public Property Let value舒张压(ByVal vNewValue As String)
    txtInfo(I舒张压).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value舒张压"
End Property

Public Property Get value血压单位() As String
    value血压单位 = cboBpUnit.Text
End Property

Public Property Let value血压单位(ByVal vNewValue As String)
    Call CboLocate(cboBpUnit, vNewValue)
    PropertyChanged "value血压单位"
End Property

Private Sub SetLine(ByVal lngStyle As Long)
'功能：设置下划线
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    UserControl.Cls
    If lngStyle = Underline Then
        For i = 0 To txtInfo.Count - 1
            x1 = txtInfo(i).Left
            y1 = txtInfo(i).Top + txtInfo(i).Height
            x2 = txtInfo(i).Left + txtInfo(i).Width
            y2 = y1
            UserControl.Line (x1, y1)-(x2, y2), &H808080
        Next
    
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
        Call gobjCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call gobjControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strMask As String
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call gobjCommFun.PressKey(vbKeyTab)
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        
        Select Case Index
            Case I呼吸, I脉搏
                strMask = "1234567890"
            Case I收缩压, I舒张压
                If cboBpUnit.Text = "mmHg" Then
                    strMask = "1234567890"
                Else
                    strMask = "1234567890."
                End If
            Case I体温
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
    Me.HaveColon = PropBag.ReadProperty("HaveColon", False)
    
    Me.value体温 = PropBag.ReadProperty("value体温", "")
    Me.value脉搏 = PropBag.ReadProperty("value脉搏", "")
    Me.value呼吸 = PropBag.ReadProperty("value呼吸", "")
    Me.value收缩压 = PropBag.ReadProperty("value收缩压", "")
    Me.value舒张压 = PropBag.ReadProperty("value舒张压", "")
    Me.value血压单位 = PropBag.ReadProperty("value血压单位", "")
    Me.Tag = PropBag.ReadProperty("Tag", "")
    
    UserControl_Resize
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
    
    For i = I体温 To I收缩压
        If i = I呼吸 And mEnumShowMode = TwoRow Then
            lngTop = txtInfo(0).Height + Me.YDis
            lngLeft = 0
        End If
        lblName(i).Move lngLeft, lngTop, UserControl.TextWidth(IIf(Me.Style = Underline, "中联:", "中联")), lngHeight
        lngLeft = lngLeft + lblName(i).Width + Me.LabToTxt
        txtInfo(i).Move lngLeft, lngTop, UserControl.TextWidth("中联信"), lngHeight
        lngLeft = lngLeft + txtInfo(i).Width
        lblUnit(i).Move lngLeft, lngTop, IIf(mEnumShowMode = OneRow, lblUnit(i).Width, UserControl.TextWidth("中/联")), lngHeight
        If mEnumShowMode = OneRow Then
            lblUnit(i).Width = UserControl.TextWidth(lblUnit(i).Caption)
        End If
        lngLeft = lngLeft + lblUnit(i).Width + Me.XDis
    Next
    
    If mblnColon Then
        For i = I体温 To I收缩压
            If InStr(lblName(i).Caption, ":") = 0 Then
                lblName(i).Caption = lblName(i).Caption & ":"
            Else
                lblName(i).Caption = Replace(lblName(i).Caption, ":", "")
            End If
        Next
    End If
    
    lblName(I舒张压).Top = lngTop
    lblName(I舒张压).Width = UserControl.TextWidth("/")
    lblName(I舒张压).Height = lngHeight
    lblName(I舒张压).Left = txtInfo(I收缩压).Left + txtInfo(I收缩压).Width
    
    txtInfo(I舒张压).Move lblName(I舒张压).Left + lblName(I舒张压).Width, lngTop, UserControl.TextWidth("中联信"), lngHeight
    
    If Me.Style = TextBox Then
        txtInfo(I舒张压).Height = IIf(txtInfo(I舒张压).Height < 300, 300, txtInfo(I舒张压).Height)
        For i = I体温 To I收缩压
            txtInfo(i).Height = IIf(txtInfo(i).Height < 300, 300, txtInfo(i).Height)
            lblName(i).Top = (txtInfo(i).Height - lblName(i).Height) / 2 + txtInfo(i).Top
            lblUnit(i).Top = (txtInfo(i).Height - lblUnit(i).Height) / 2 + txtInfo(i).Top
        Next
        lblName(I舒张压).Top = (txtInfo(I舒张压).Height - lblName(I舒张压).Height) / 2 + txtInfo(I舒张压).Top
    End If
    
    cboBpUnit.Left = txtInfo(I舒张压).Left + txtInfo(I舒张压).Width
    cboBpUnit.Top = txtInfo(I舒张压).Top
    cboBpUnit.Width = UserControl.TextWidth("mmHg") + 400
    cboBpUnit.Height = lngHeight * 2
    
    cboBpUnit.Left = IIf(Me.Style = Underline, -30, 0)
    cboBpUnit.Top = IIf(Me.Style = Underline, -30, 0)
    
    If Me.Style = Underline Then
        fraCboBorder.Height = IIf(cboBpUnit.Height <= 300, 240, cboBpUnit.Height - 40)
        fraCboBorder.Top = txtInfo(I舒张压).Top
        fraCboBorder.Width = UserControl.TextWidth("mmHg") + 350
    Else
        fraCboBorder.Height = txtInfo(0).Height
        fraCboBorder.Top = txtInfo(I舒张压).Top
        fraCboBorder.Width = cboBpUnit.Width
    End If
    
    fraCboBorder.Left = txtInfo(I舒张压).Left + txtInfo(I舒张压).Width
    
    UserControl.Width = fraCboBorder.Left + fraCboBorder.Width
    UserControl.Height = txtInfo(I舒张压).Top + txtInfo(I舒张压).Height
End Sub

Private Sub UserControl_Show()
    Call SetLine(Me.Style)
End Sub

Private Sub UserControl_Terminate()
    mblnSaveNow = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ControlLock", txtInfo(I体温).locked, False)
    Call PropBag.WriteProperty("Enabled", Me.Enabled, True)
    Call PropBag.WriteProperty("TextBackColor", txtInfo(I体温).BackColor, &H80000000)
    Call PropBag.WriteProperty("LblBackColor", lblName(I体温).BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", Me.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", mcolForeColor, &H80000000)
    Call PropBag.WriteProperty("BackColor", Me.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ShowMode", Me.ShowMode, enum_ShowMode.TwoRow)
    Call PropBag.WriteProperty("Style", Me.Style, enum_Style.TextBox)
    Call PropBag.WriteProperty("MaxLength", txtInfo(I体温).MaxLength, 0)
    Call PropBag.WriteProperty("XDis", mXDis, 20)
    Call PropBag.WriteProperty("YDis", mYDis, 10)
    Call PropBag.WriteProperty("LabToTxt", mLabToTxt, 10)
    Call PropBag.WriteProperty("HaveColon", mblnColon, False)
    
    Call PropBag.WriteProperty("value呼吸", Me.value呼吸, "")
    Call PropBag.WriteProperty("value脉搏", Me.value脉搏, "")
    Call PropBag.WriteProperty("value体温", Me.value体温, "")
    Call PropBag.WriteProperty("value收缩压", Me.value收缩压, "")
    Call PropBag.WriteProperty("value舒张压", Me.value舒张压, "")
    Call PropBag.WriteProperty("value血压单位", Me.value血压单位, "")
    Call PropBag.WriteProperty("Tag", Me.Tag, "")
End Sub

Public Function LoadPatiVitalSigns(ByVal lng收发ID As Long, ByVal int性质 As Integer)
'功能：加载记录内容到对应的文本框中，主要将数据保存于文本框的Tag值中
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Integer
    
    mblnNoCheck = True
    
    mlng收发ID = lng收发ID
    mint性质 = int性质
    
    Call ClearData
    
    With cboBpUnit
        .Clear
        .AddItem "mmHg"
        .AddItem "Kpa"
    End With
    
    Set mcol范围 = New Collection
    
    strSQL = "Select ID, 中文名, 长度, 单位, 数值域, 小数 From 诊治所见项目" & _
        " Where 分类id = 7 And 中文名 In ('体温', '脉搏', '收缩压', '舒张压', '呼吸')"
        
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "PatVitalSigns")
    
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
        
            strTmp = "" & rsTmp!数值域
            If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
            mcol范围.Add strTmp, "_" & rsTmp!id
            strTmp = Replace(strTmp, ";", " - ")
            If Val("" & rsTmp!小数) = 0 Then
                strTmp = "范围为 " & strTmp & " 的数。"
            Else
                strTmp = "范围为 " & strTmp & " 之间的数，可含" & rsTmp!小数 & "位小数，最多可输入" & rsTmp!长度 & "个字符。"
            End If
            Select Case rsTmp!中文名
                Case "体温"
                    lblName(I体温).Tag = rsTmp!id
                    lblUnit(I体温).Caption = Nvl(rsTmp!单位, "℃")
                    txtInfo(I体温).MaxLength = 4
                    txtInfo(I体温).ToolTipText = "体温" & strTmp
                Case "呼吸"
                    lblName(I呼吸).Tag = rsTmp!id
                    lblUnit(I呼吸).Caption = Nvl(rsTmp!单位, "次/分")
                    txtInfo(I呼吸).MaxLength = 3
                    txtInfo(I呼吸).ToolTipText = "呼吸" & strTmp
                Case "脉搏"
                    lblName(I脉搏).Tag = rsTmp!id
                    lblUnit(I脉搏).Caption = Nvl(rsTmp!单位, "次/分")
                    txtInfo(I脉搏).MaxLength = 3
                    txtInfo(I脉搏).ToolTipText = "脉搏" & strTmp
                Case "收缩压"
                    lblName(I收缩压).Tag = rsTmp!id
                    Call gobjControl.cbo.SetIndex(cboBpUnit.hWnd, IIf("" & rsTmp!单位 = "mmHg", 0, 1))
                    txtInfo(I收缩压).MaxLength = 5
                    txtInfo(I收缩压).ToolTipText = "收缩压" & strTmp
                Case "舒张压"
                    lblName(I舒张压).Tag = rsTmp!id
                    Call gobjControl.cbo.SetIndex(cboBpUnit.hWnd, IIf("" & rsTmp!单位 = "mmHg", 0, 1))
                    txtInfo(I舒张压).MaxLength = 5
                    txtInfo(I舒张压).ToolTipText = "舒张压" & strTmp
            End Select
            rsTmp.MoveNext
        Next
    End If
    cboBpUnit.Tag = cboBpUnit.List(cboBpUnit.ListIndex)
    
    strSQL = "select 体温,脉搏,呼吸,收缩压,舒张压,血压单位 from 血液执行生命体征 where 收发ID=[1] and 性质=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "PatVitalSigns", mlng收发ID, mint性质)
        
    If rsTmp.RecordCount > 0 Then
        For i = 0 To rsTmp.Fields.Count - 1
            Select Case rsTmp.Fields(i).name
                Case "体温"
                    txtInfo(I体温).Text = "" & rsTmp.Fields(i).Value
                    txtInfo(I体温).Tag = "" & rsTmp.Fields(i).Value
                Case "呼吸"
                    txtInfo(I呼吸).Text = "" & rsTmp.Fields(i).Value
                    txtInfo(I呼吸).Tag = "" & rsTmp.Fields(i).Value
                Case "脉搏"
                    txtInfo(I脉搏).Text = "" & rsTmp.Fields(i).Value
                    txtInfo(I脉搏).Tag = "" & rsTmp.Fields(i).Value
                Case "收缩压"
                    txtInfo(I收缩压).Text = "" & rsTmp.Fields(i).Value
                    txtInfo(I收缩压).Tag = "" & rsTmp.Fields(i).Value
                Case "舒张压"
                    txtInfo(I舒张压).Text = "" & rsTmp.Fields(i).Value
                    txtInfo(I舒张压).Tag = "" & rsTmp.Fields(i).Value
                Case "血压单位"
                    If Not IsNull(rsTmp.Fields(i).Value) Then Call gobjControl.cbo.SetIndex(cboBpUnit.hWnd, IIf("" & rsTmp.Fields(i).Value = "mmHg", 0, 1))
            End Select
        Next
    End If
    If cboBpUnit.ListIndex = -1 Then Call gobjControl.cbo.SetIndex(cboBpUnit, 0)
    Call BpRange(cboBpUnit.List(cboBpUnit.ListIndex))
    cboBpUnit.Tag = cboBpUnit.List(cboBpUnit.ListIndex)
    mblnNoCheck = False
    mstrPreState = InSideSaveSQL
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Function

Public Function GiveUpSave()
'功能：放弃修改值不保存，则恢复原值
    Dim i As Integer
    For i = 0 To 6
        txtInfo(i).Text = txtInfo(i).Tag
    Next
End Function

Public Function GetSaveSQL(ByVal lng收发ID As Long, ByVal int性质 As Integer) As String
'功能：返回门诊生命体征填写的SQL
    Dim strTmp As String
    Dim strSQL As String
    Dim i As Integer
    For i = I体温 To I舒张压
        If IsNumeric(txtInfo(i).Text) Then
            strTmp = gobjComlib.FormatEx(Val(txtInfo(i).Text), 2)
        Else
            strTmp = "NULL"
        End If
        strSQL = strSQL & "," & strTmp
    Next
    GetSaveSQL = "Zl_血液执行生命体征_Update(" & lng收发ID & "," & int性质 & strSQL & ",'" & cboBpUnit.Text & "')"
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
            For i = I收缩压 To I舒张压
                strTmp = GetCollectContent(mcol范围, "_" & lblName(i).Tag)
                If InStr(strTmp, ";") > 0 Then
                    dblMin = Val(Split(strTmp, ";")(0))
                    dblMax = Val(Split(strTmp, ";")(1))
                    
                    dblMin = Round(dblMin * 10 * 3 / 4)
                    dblMax = Round(dblMax * 10 * 3 / 4)
                    txtInfo(i).ToolTipText = IIf(i = I收缩压, "收缩", "舒张") & "压范围为 " & dblMin & " - " & dblMax & str单位
                    mcol范围.Remove ("_" & lblName(i).Tag)
                    mcol范围.Add dblMin & ";" & dblMax, "_" & lblName(i).Tag
                    txtInfo(i).MaxLength = 3
                End If
            Next
        Else
            For i = I收缩压 To I舒张压
                strTmp = GetCollectContent(mcol范围, "_" & lblName(i).Tag)
                If InStr(strTmp, ";") > 0 Then
                    dblMin = Val(Split(strTmp, ";")(0))
                    dblMax = Val(Split(strTmp, ";")(1))
                    
                    dblMin = Format(dblMin * 4 / 3 / 10, "#0.00")
                    dblMax = Format(dblMax * 4 / 3 / 10, "#0.00")
                    txtInfo(i).ToolTipText = IIf(i = I收缩压, "收缩", "舒张") & "压范围为 " & dblMin & " - " & dblMax & str单位
                    mcol范围.Remove ("_" & lblName(i).Tag)
                    mcol范围.Add dblMin & ";" & dblMax, "_" & lblName(i).Tag
                    txtInfo(i).MaxLength = 5
                End If
            Next
        End If
    End If
End Sub

Private Sub cboBpUnit_Click()
'血压单位换算
    Dim strTmp As String
    
    If cboBpUnit.List(cboBpUnit.ListIndex) <> cboBpUnit.Tag Then
        If cboBpUnit.List(cboBpUnit.ListIndex) = "mmHg" Then
            'Kpa转换到mmHg 乘10再乘3减半再减半(mmHg四舍五入)
            If txtInfo(I舒张压).Text <> "" Then
                txtInfo(I舒张压).Text = Round(Val(txtInfo(I舒张压).Text) * 10 * 3 / 4)
            End If
            If txtInfo(I收缩压).Text <> "" Then
                txtInfo(I收缩压).Text = Round(Val(txtInfo(I收缩压).Text) * 10 * 3 / 4)
            End If
        Else
            'mmHg转换到Kpa 加倍加倍除3再除10(Kpa保留两位小数)
            If txtInfo(I舒张压).Text <> "" Then
                txtInfo(I舒张压).Text = Format(Val(txtInfo(I舒张压).Text) * 4 / 3 / 10, "#0.00")
            End If
            If txtInfo(I收缩压).Text <> "" Then
                txtInfo(I收缩压).Text = Format(Val(txtInfo(I收缩压).Text) * 4 / 3 / 10, "#0.00")
            End If
        End If
        Call BpRange(cboBpUnit.List(cboBpUnit.ListIndex))
        cboBpUnit.Tag = cboBpUnit.List(cboBpUnit.ListIndex)
        RaiseEvent Change(I舒张压 + 1)
        
        If Not mblnNoCheck And mblnSaveNow Then
            strTmp = InSideSaveSQL
            If strTmp <> mstrPreState Then
                Call gobjDatabase.ExecuteProcedure(strTmp, "生命体征")
                mstrPreState = strTmp
            End If
        End If
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If txtInfo(Index).Text = txtInfo(Index).Tag Then Exit Sub
    RaiseEvent Change(Index)
End Sub

Private Sub cboBpUnit_Change()
    RaiseEvent Change(I舒张压 + 1)
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
'判断范围值
    Dim strTmp As String
    Dim dblMin As Double
    Dim dblMax As Double
    
    If txtInfo(Index).Text <> "" Then
        If Not IsNumeric(txtInfo(Index).Text) Then
            MsgBox "输入内容必须是数字，" & txtInfo(Index).ToolTipText, vbInformation, gstrSysName
            txtInfo(Index).Text = txtInfo(Index).Tag
            Cancel = True
            Call gobjControl.TxtSelAll(txtInfo(Index))
            Exit Sub
        End If
        
        strTmp = GetCollectContent(mcol范围, "_" & lblName(Index).Tag)
        If InStr(strTmp, ";") > 0 Then
            dblMin = Val(Split(strTmp, ";")(0))
            dblMax = Val(Split(strTmp, ";")(1))
            
            If Val(txtInfo(Index).Text) > dblMax Or Val(txtInfo(Index).Text) < dblMin Then
                MsgBox "输入内容未在指定范围内，" & txtInfo(Index).ToolTipText, vbInformation, gstrSysName
                txtInfo(Index).Text = txtInfo(Index).Tag
                Cancel = True
                Call gobjControl.TxtSelAll(txtInfo(Index))
                Exit Sub
            End If
        End If
    End If
    If Not mblnNoCheck And mblnSaveNow Then
        strTmp = InSideSaveSQL
        If strTmp <> mstrPreState Then
            Call gobjDatabase.ExecuteProcedure(strTmp, "生命体征")
            mstrPreState = strTmp
        End If
    End If
End Sub

Public Function ClearTxtToolTipText()
'功能：清空所有文本框的提示
    Dim i As Integer
    
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).ToolTipText = ""
    Next
End Function

Public Sub TxtAlignment(ByVal intType As Integer)
'功能：设置文本框的对齐方式
'intType 0-左对齐，1－右对齐，2－居中
    Dim i As Long
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).Alignment = intType
    Next
End Sub

Public Sub SetUseType(ByVal blnSaveNow As Boolean, Optional ByVal strTag As String)
    mblnSaveNow = blnSaveNow
End Sub

Private Function InSideSaveSQL() As String
'功能：返回门诊生命体征填写的SQL
    Dim strTmp As String
    Dim strSQL As String
    Dim i As Integer
    If Not mblnSaveNow Then Exit Function
    For i = I体温 To I舒张压
        If IsNumeric(txtInfo(i).Text) Then
            strTmp = gobjComlib.FormatEx(Val(txtInfo(i).Text), 2)
        Else
            strTmp = "NULL"
        End If
        strSQL = strSQL & "," & strTmp
    Next
    InSideSaveSQL = "Zl_血液执行生命体征_Update(" & mlng收发ID & "," & mint性质 & strSQL & ",'" & cboBpUnit.Text & "')"
End Function

Private Function GetCollectContent(ByVal objCollect As Collection, ByVal strKey As String) As String
    Dim strRetrun As String
    On Error Resume Next
    Err.Clear
    strRetrun = objCollect(strKey)
    If Err <> 0 Then
        Err.Clear
        strRetrun = ""
    End If
    GetCollectContent = strRetrun
End Function

