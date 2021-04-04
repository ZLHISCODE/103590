VERSION 5.00
Begin VB.UserControl UCPatiVitalSigns 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6165
   ScaleHeight     =   750
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
      Left            =   1950
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5295
      TabIndex        =   0
      Top             =   405
      Width           =   765
      Begin VB.ComboBox cboBpUnit 
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
      AutoSize        =   -1  'True
      Caption         =   "���"
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   15
      Width           =   360
   End
   Begin VB.Label lblName 
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   1560
      TabIndex        =   19
      Top             =   60
      Width           =   400
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   3615
      TabIndex        =   18
      Top             =   60
      Width           =   360
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   3
      Left            =   30
      TabIndex        =   17
      Top             =   450
      Width           =   360
   End
   Begin VB.Label lblName 
      Caption         =   "����"
      Height          =   180
      Index           =   4
      Left            =   1560
      TabIndex        =   16
      Top             =   405
      Width           =   405
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Ѫѹ"
      Height          =   180
      Index           =   5
      Left            =   3570
      TabIndex        =   15
      Top             =   390
      Width           =   360
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
      AutoSize        =   -1  'True
      Caption         =   "cm"
      Height          =   180
      Index           =   0
      Left            =   1095
      TabIndex        =   13
      Top             =   60
      Width           =   180
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "kg"
      Height          =   180
      Index           =   1
      Left            =   2730
      TabIndex        =   12
      Top             =   75
      Width           =   180
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Index           =   2
      Left            =   4890
      TabIndex        =   11
      Top             =   60
      Width           =   180
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "��/��"
      Height          =   180
      Index           =   3
      Left            =   1065
      TabIndex        =   10
      Top             =   420
      Width           =   450
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "��/��"
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

Public Event Change(ByVal int��� As Integer)
'    I��� = 0
'    I���� = 1
'    I���� = 2
'    I���� = 3
'    I���� = 4
'    I����ѹ = 5
'    I����ѹ = 6
'    Ѫѹ��λ = 7

Private Enum E_ITEM_INDEX
    I��� = 0
    I���� = 1
    I���� = 2
    I���� = 3
    I���� = 4
    
    I����ѹ = 5
    I����ѹ = 6
    
    IѪѹ�ָ��� = 6
End Enum
 
Public Enum enum_Style '�ı���ķ�� Ĭ��0-TextBox
    TextBox = 0
    Underline = 1
End Enum
Private mEnumStyle As enum_Style

Public Enum enum_ShowMode '������Ŀǰֻ�ṩ1�к�2����ʽ
    OneRow = 0
    TwoRow = 1
End Enum
Private mEnumShowMode As enum_ShowMode

Private mXDis As Long 'ˮƽ�������Ŀ֮��ļ��
Private mYDis As Long '��ֱ�������Ŀ֮��ļ��������������ʾģʽ����Ч
Private mLabToTxt As Long '��ǩ�����ı���ľ���
Private mcolForeColor As OLE_COLOR
Private mstrTag As String
Private mcol��Χ As Collection '�����ȡֵ��Χ
Private mblnColon As Boolean '��ǩ�����Ƿ���ð��

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Get Hwnd() As Long
    Hwnd = UserControl.Hwnd
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
    ControlLock = txtInfo(I���).Locked
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
    TextBackColor = txtInfo(I���).BackColor
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
    LblBackColor = lblName(I���).BackColor
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
    MaxLength = txtInfo(I���).MaxLength
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

Public Property Get value���() As String
    value��� = txtInfo(I���).Text
End Property

Public Property Let value���(ByVal vNewValue As String)
    txtInfo(I���).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value���"
End Property

Public Property Get value����() As String
    value���� = txtInfo(I����).Text
End Property

Public Property Let value����(ByVal vNewValue As String)
    txtInfo(I����).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����"
End Property

Public Property Get value����() As String
    value���� = txtInfo(I����).Text
End Property

Public Property Let value����(ByVal vNewValue As String)
    txtInfo(I����).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����"
End Property

Public Property Get value����() As String
    value���� = txtInfo(I����).Text
End Property

Public Property Let value����(ByVal vNewValue As String)
    txtInfo(I����).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����"
End Property

Public Property Get value����() As String
    value���� = txtInfo(I����).Text
End Property

Public Property Let value����(ByVal vNewValue As String)
    txtInfo(I����).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����"
End Property

Public Property Get value����ѹ() As String
    value����ѹ = txtInfo(I����ѹ).Text
End Property

Public Property Let value����ѹ(ByVal vNewValue As String)
    txtInfo(I����ѹ).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����ѹ"
End Property

Public Property Get value����ѹ() As String
    value����ѹ = txtInfo(I����ѹ).Text
End Property

Public Property Let value����ѹ(ByVal vNewValue As String)
    txtInfo(I����ѹ).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����ѹ"
End Property

Public Property Get valueѪѹ��λ() As String
    valueѪѹ��λ = cboBpUnit.Text
End Property

Public Property Let valueѪѹ��λ(ByVal vNewValue As String)
    Call cbo.Locate(cboBpUnit, vNewValue)
    PropertyChanged "valueѪѹ��λ"
End Property

Private Sub SetLine(ByVal lngStyle As Long)
'���ܣ������»���
    Dim i As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
    
    UserControl.Cls
    If lngStyle = Underline Then
        For i = 0 To txtInfo.Count - 1
            X1 = txtInfo(i).Left
            Y1 = txtInfo(i).Top + txtInfo(i).Height
            X2 = txtInfo(i).Left + txtInfo(i).Width
            Y2 = Y1
            UserControl.Line (X1, Y1)-(X2, Y2), &H808080
        Next
    
        X1 = fraCboBorder.Left
        Y1 = fraCboBorder.Top + fraCboBorder.Height
        X2 = fraCboBorder.Left + fraCboBorder.Width
        Y2 = Y1
        UserControl.Line (X1, Y1)-(X2, Y2), &H808080
    End If
End Sub

Private Sub cboBpUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtInfo_GotFocus(index As Integer)
    Call zlControl.TxtSelAll(txtInfo(index))
End Sub

Private Sub txtInfo_KeyPress(index As Integer, KeyAscii As Integer)
    Dim strMask As String
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        
        Select Case index
            Case I���, I����, I����
                strMask = "1234567890"
                If index = I��� Then strMask = strMask & "."
            Case I����ѹ, I����ѹ
                If cboBpUnit.Text = "mmHg" Then
                    strMask = "1234567890"
                Else
                    strMask = "1234567890."
                End If
            Case I����, I����
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
    
    Me.value��� = PropBag.ReadProperty("value���", "")
    Me.value���� = PropBag.ReadProperty("value����", "")
    Me.value���� = PropBag.ReadProperty("value����", "")
    Me.value���� = PropBag.ReadProperty("value����", "")
    Me.value���� = PropBag.ReadProperty("value����", "")
    Me.value����ѹ = PropBag.ReadProperty("value����ѹ", "")
    Me.value����ѹ = PropBag.ReadProperty("value����ѹ", "")
    Me.valueѪѹ��λ = PropBag.ReadProperty("valueѪѹ��λ", "")
    Me.Tag = PropBag.ReadProperty("Tag", "")
End Sub

Private Sub UserControl_Resize()
'����: ���ÿؼ���Сλ��
    On Error Resume Next
    Dim lngHeight As Long
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim i As Integer
    
    'ȡ��������Ӧ����
    lngHeight = UserControl.TextHeight("��")
    
    lngTop = 0
    lngLeft = 0
    
    For i = 0 To 5
        If i = 3 And mEnumShowMode = TwoRow Then
        
            lngTop = txtInfo(0).Height + Me.YDis
            lngLeft = 0
        End If
        lblName(i).Move lngLeft, lngTop, UserControl.TextWidth(IIf(Me.Style = Underline, "����:", "����")), lngHeight
        lngLeft = lngLeft + lblName(i).Width + IIf(Me.Style = TextBox, Me.LabToTxt, 0)
        txtInfo(i).Move lngLeft, lngTop, UserControl.TextWidth("������"), lngHeight
        lngLeft = lngLeft + txtInfo(i).Width
        lblUnit(i).Move lngLeft, lngTop, UserControl.TextWidth("��/��"), lngHeight
        lngLeft = lngLeft + lblUnit(i).Width + Me.XDis
    Next
    
    If mblnColon Then
        For i = 0 To 5
            If InStr(lblName(i).Caption, ":") = 0 Then
                lblName(i).Caption = lblName(i).Caption & ":"
            Else
                lblName(i).Caption = Replace(lblName(i).Caption, ":", "")
            End If
        Next
    End If
    
    lblName(IѪѹ�ָ���).Top = lngTop
    lblName(IѪѹ�ָ���).Width = UserControl.TextWidth("/")
    lblName(IѪѹ�ָ���).Height = lngHeight
    lblName(IѪѹ�ָ���).Left = txtInfo(I����ѹ).Left + txtInfo(I����ѹ).Width
    
    txtInfo(I����ѹ).Move lblName(IѪѹ�ָ���).Left + lblName(IѪѹ�ָ���).Width, lngTop, UserControl.TextWidth("������"), lngHeight
    
    If Me.Style = TextBox Then
        txtInfo(I����ѹ).Height = IIf(txtInfo(I����ѹ).Height < 300, 300, txtInfo(I����ѹ).Height)
        For i = 0 To 5
            txtInfo(i).Height = IIf(txtInfo(i).Height < 300, 300, txtInfo(i).Height)
            lblName(i).Top = (txtInfo(i).Height - lblName(i).Height) / 2 + txtInfo(i).Top
            lblUnit(i).Top = (txtInfo(i).Height - lblUnit(i).Height) / 2 + txtInfo(i).Top
        Next
        lblName(IѪѹ�ָ���).Top = (txtInfo(I����ѹ).Height - lblName(IѪѹ�ָ���).Height) / 2 + txtInfo(I����ѹ).Top
    End If
    
    cboBpUnit.Left = txtInfo(I����ѹ).Left + txtInfo(I����ѹ).Width
    cboBpUnit.Top = txtInfo(I����ѹ).Top
    cboBpUnit.Width = UserControl.TextWidth("mmHg") + 400
    cboBpUnit.Height = lngHeight * 2
    
    cboBpUnit.Left = IIf(Me.Style = Underline, -30, 0)
    cboBpUnit.Top = IIf(Me.Style = Underline, -30, 0)
    
    If Me.Style = Underline Then
        fraCboBorder.Height = IIf(cboBpUnit.Height <= 300, 240, cboBpUnit.Height - 40)
        fraCboBorder.Top = txtInfo(I����ѹ).Top
        fraCboBorder.Width = UserControl.TextWidth("mmHg") + 350
    Else
        fraCboBorder.Height = txtInfo(0).Height
        fraCboBorder.Top = txtInfo(I����ѹ).Top
        fraCboBorder.Width = cboBpUnit.Width
    End If
    
    fraCboBorder.Left = txtInfo(I����ѹ).Left + txtInfo(I����ѹ).Width
    
    UserControl.Width = fraCboBorder.Left + fraCboBorder.Width
    UserControl.Height = txtInfo(I����ѹ).Top + txtInfo(I����ѹ).Height + 100
End Sub

Private Sub UserControl_Show()
    Call SetLine(Me.Style)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ControlLock", txtInfo(I���).Locked, False)
    Call PropBag.WriteProperty("Enabled", Me.Enabled, True)
    Call PropBag.WriteProperty("TextBackColor", txtInfo(I���).BackColor, &H80000000)
    Call PropBag.WriteProperty("LblBackColor", lblName(I���).BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", Me.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", mcolForeColor, &H80000000)
    Call PropBag.WriteProperty("BackColor", Me.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ShowMode", Me.ShowMode, enum_ShowMode.TwoRow)
    Call PropBag.WriteProperty("Style", Me.Style, enum_Style.TextBox)
    Call PropBag.WriteProperty("MaxLength", txtInfo(I���).MaxLength, 0)
    Call PropBag.WriteProperty("XDis", mXDis, 20)
    Call PropBag.WriteProperty("YDis", mYDis, 10)
    Call PropBag.WriteProperty("LabToTxt", mLabToTxt, 10)
    Call PropBag.WriteProperty("HaveColon", mblnColon, False)
    
    Call PropBag.WriteProperty("value���", Me.value���, "")
    Call PropBag.WriteProperty("value����", Me.value����, "")
    Call PropBag.WriteProperty("value����", Me.value����, "")
    Call PropBag.WriteProperty("value����", Me.value����, "")
    Call PropBag.WriteProperty("value����", Me.value����, "")
    Call PropBag.WriteProperty("value����ѹ", Me.value����ѹ, "")
    Call PropBag.WriteProperty("value����ѹ", Me.value����ѹ, "")
    Call PropBag.WriteProperty("valueѪѹ��λ", Me.valueѪѹ��λ, "")
    Call PropBag.WriteProperty("Tag", Me.Tag, "")
End Sub

Public Function LoadPatiVitalSigns(ByVal lng����ID As Long, ByVal lng�Һ�id As Long)
'���ܣ����ؼ�¼���ݵ���Ӧ���ı����У���Ҫ�����ݱ������ı����Tagֵ��
'ע�⣺���غ�Ѫѹ�ȵ�����ʱ�ݲ��ϸ���ƣ������Ա�֤������ڹ涨�ķ�Χ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Integer
    
    Call ClearData
    
    With cboBpUnit
        .Clear
        .AddItem "mmHg"
        .AddItem "Kpa"
    End With
    
    Set mcol��Χ = New Collection
    
    strSQL = "Select ID, ������, ����, ��λ, ��ֵ��, С�� From ����������Ŀ" & _
        " Where ����id = 7 And ������ In ('����', '����', '����ѹ', '����ѹ', '����', '���', '����')"
        
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatVitalSigns")
    
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
        
            strTmp = "" & rsTmp!��ֵ��
            If Left(strTmp, 1) = "." Then strTmp = "0" & strTmp
            mcol��Χ.Add strTmp, "_" & rsTmp!id
            strTmp = Replace(strTmp, ";", " - ")
            If Val("" & rsTmp!С��) = 0 Then
                strTmp = "��ΧΪ " & strTmp & " ������"
            Else
                strTmp = "��ΧΪ " & strTmp & " ֮��������ɺ�" & rsTmp!С�� & "λС������������" & rsTmp!���� & "���ַ���"
            End If
            Select Case rsTmp!������
                Case "���"
                    lblName(I���).Tag = rsTmp!id
                    lblUnit(I���).Caption = rsTmp!��λ
                    txtInfo(I���).MaxLength = rsTmp!����
                    txtInfo(I���).ToolTipText = "���" & strTmp
                Case "����"
                    lblName(I����).Tag = rsTmp!id
                    lblUnit(I����).Caption = rsTmp!��λ
                    txtInfo(I����).MaxLength = 5 'rsTmp!����
                    txtInfo(I����).ToolTipText = "����" & strTmp
                Case "����"
                    lblName(I����).Tag = rsTmp!id
                    lblUnit(I����).Caption = rsTmp!��λ
                    txtInfo(I����).MaxLength = rsTmp!����
                    txtInfo(I����).ToolTipText = "����" & strTmp
                Case "����"
                    lblName(I����).Tag = rsTmp!id
                    lblUnit(I����).Caption = rsTmp!��λ
                    txtInfo(I����).MaxLength = rsTmp!����
                    txtInfo(I����).ToolTipText = "����" & strTmp
                Case "����"
                    lblName(I����).Tag = rsTmp!id
                    lblUnit(I����).Caption = rsTmp!��λ
                    txtInfo(I����).MaxLength = rsTmp!����
                    txtInfo(I����).ToolTipText = "����" & strTmp
                Case "����ѹ"
                    lblName(I����ѹ).Tag = rsTmp!id
                    Call cbo.SetIndex(cboBpUnit.Hwnd, IIf("" & rsTmp!��λ = "mmHg", 0, 1))
                    txtInfo(I����ѹ).MaxLength = 5 'rsTmp!����
                    txtInfo(I����ѹ).ToolTipText = "����ѹ" & strTmp
                Case "����ѹ"
                    lblName(IѪѹ�ָ���).Tag = rsTmp!id
                    Call cbo.SetIndex(cboBpUnit.Hwnd, IIf("" & rsTmp!��λ = "mmHg", 0, 1))
                    txtInfo(I����ѹ).MaxLength = 5 'rsTmp!����
                    txtInfo(I����ѹ).ToolTipText = "����ѹ" & strTmp
            End Select
            rsTmp.MoveNext
        Next
    End If
    cboBpUnit.Tag = cboBpUnit.List(cboBpUnit.ListIndex)
    
    strSQL = "Select b.��Ŀ��λ, b.��Ŀ����, b.��¼����" & _
        " From ���˻����¼ A, ���˻������� B Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatVitalSigns", lng����ID, lng�Һ�id)
        
    If rsTmp.RecordCount <= 0 Then
        strSQL = "Select '' as ��Ŀ��λ, ��Ϣ�� As ��Ŀ����, ��Ϣֵ As ��¼���� From ������Ϣ�ӱ� Where ����id = [1] And ����id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatVitalSigns", lng����ID, lng�Һ�id)
    End If
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            Select Case rsTmp!��Ŀ����
                Case "���"
                    txtInfo(I���).Text = rsTmp!��¼����
                    txtInfo(I���).Tag = rsTmp!��¼����
                Case "����"
                    txtInfo(I����).Text = rsTmp!��¼����
                    txtInfo(I����).Tag = rsTmp!��¼����
                Case "����"
                    txtInfo(I����).Text = rsTmp!��¼����
                    txtInfo(I����).Tag = rsTmp!��¼����
                Case "����"
                    txtInfo(I����).Text = rsTmp!��¼����
                    txtInfo(I����).Tag = rsTmp!��¼����
                Case "����"
                    txtInfo(I����).Text = rsTmp!��¼����
                    txtInfo(I����).Tag = rsTmp!��¼����
                Case "����ѹ"
                    txtInfo(I����ѹ).Text = rsTmp!��¼����
                    txtInfo(I����ѹ).Tag = rsTmp!��¼����
                    Call cbo.SetIndex(cboBpUnit.Hwnd, IIf("" & rsTmp!��Ŀ��λ = "mmHg", 0, 1))
                Case "����ѹ"
                    txtInfo(I����ѹ).Text = rsTmp!��¼����
                    txtInfo(I����ѹ).Tag = rsTmp!��¼����
                    Call cbo.SetIndex(cboBpUnit.Hwnd, IIf("" & rsTmp!��Ŀ��λ = "mmHg", 0, 1))
            End Select
            rsTmp.MoveNext
        Next
    End If
    If cboBpUnit.ListIndex = -1 Then Call cbo.SetIndex(cboBpUnit, 0)
    Call BpRange(cboBpUnit.List(cboBpUnit.ListIndex))
    cboBpUnit.Tag = cboBpUnit.List(cboBpUnit.ListIndex)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GiveUpSave()
'���ܣ������޸�ֵ�����棬��ָ�ԭֵ
    Dim i As Integer
    For i = 0 To 6
        txtInfo(i).Text = txtInfo(i).Tag
    Next
End Function
Public Function GetSaveSQL(ByVal lng����ID As Long, ByVal lng�Һ�id As Long) As String
'���ܣ�������������������д��SQL
    GetSaveSQL = GetRetrunSQL(lng����ID, lng�Һ�id)
End Function

Public Function GetRetrunSQL(Optional lng����ID As Long, Optional lng�Һ�id As Long) As String
'���ܣ�������������������д��SQL
    Dim strTmp As String
    Dim strSQL As String
    Dim i As Integer
    For i = 0 To 4
        If txtInfo(i).Text = "" Then
            strTmp = IIf(strTmp <> "", strTmp, "") & lblName(i).Tag & "|��|" & lblUnit(i).Caption & ","
        Else
            strTmp = IIf(strTmp <> "", strTmp, "") & lblName(i).Tag & "|" & FormatEx(Val(txtInfo(i).Text), 2) & "|" & lblUnit(i).Caption & ","
        End If
    Next
    If txtInfo(5).Text = "" Then
        strTmp = strTmp & lblName(5).Tag & "|��|" & cboBpUnit.Text & ","
    Else
        strTmp = strTmp & lblName(5).Tag & "|" & FormatEx(Val(txtInfo(5).Text), 2) & "|" & cboBpUnit.Text & ","
    End If
    If txtInfo(6).Text = "" Then
        strTmp = strTmp & lblName(6).Tag & "|��|" & cboBpUnit.Text
    Else
        strTmp = strTmp & lblName(6).Tag & "|" & FormatEx(Val(txtInfo(6).Text), 2) & "|" & cboBpUnit.Text
    End If
    GetRetrunSQL = "Zl_������������_Update(" & lng����ID & "," & lng�Һ�id & ",'" & strTmp & "')"
End Function

Public Function ClearData()
'���ܣ��л�����ʱ������ݣ���������ı����ֵ��Tagֵ
    Dim i As Integer
    
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).Text = "" '�����ֵ
        txtInfo(i).Tag = "" 'ԭֵ
    Next
    
    For i = 0 To lblName.Count - 1
        lblName(i).Tag = "" ' ��Ŀid
    Next
    
End Function

Private Sub BpRange(ByVal str��λ As String)
'���ܣ�Ѫѹ��λ�仯��ȡֵ��Χ��֮�仯 '����ѹ--5 '����ѹ--6
    Dim strTmp As String
    Dim dblMin As Double
    Dim dblMax As Double
    Dim i As Integer
    
    If cboBpUnit.Tag <> str��λ Then
        If str��λ = "mmHg" Then
            For i = 5 To 6
                strTmp = mcol��Χ("_" & lblName(i).Tag)
                If InStr(strTmp, ";") > 0 Then
                    dblMin = Val(Split(strTmp, ";")(0))
                    dblMax = Val(Split(strTmp, ";")(1))
                    
                    dblMin = Round(dblMin * 10 * 3 / 4)
                    dblMax = Round(dblMax * 10 * 3 / 4)
                    txtInfo(i).ToolTipText = IIf(i = 5, "����", "����") & "ѹ��ΧΪ " & dblMin & " - " & dblMax & str��λ
                    mcol��Χ.Remove ("_" & lblName(i).Tag)
                    mcol��Χ.Add dblMin & ";" & dblMax, "_" & lblName(i).Tag
                End If
            Next
        Else
            For i = 5 To 6
                strTmp = mcol��Χ("_" & lblName(i).Tag)
                If InStr(strTmp, ";") > 0 Then
                    dblMin = Val(Split(strTmp, ";")(0))
                    dblMax = Val(Split(strTmp, ";")(1))
                    
                    dblMin = Format(dblMin * 4 / 3 / 10, "#0.00")
                    dblMax = Format(dblMax * 4 / 3 / 10, "#0.00")
                    txtInfo(i).ToolTipText = IIf(i = 5, "����", "����") & "ѹ��ΧΪ " & dblMin & " - " & dblMax & str��λ
                    mcol��Χ.Remove ("_" & lblName(i).Tag)
                    mcol��Χ.Add dblMin & ";" & dblMax, "_" & lblName(i).Tag
                End If
            Next
        End If
    End If
End Sub

Private Sub cboBpUnit_Click()
'Ѫѹ��λ����
    If cboBpUnit.List(cboBpUnit.ListIndex) <> cboBpUnit.Tag Then
        If cboBpUnit.List(cboBpUnit.ListIndex) = "mmHg" Then
            'Kpaת����mmHg ��10�ٳ�3�����ټ���(mmHg��������)
            If txtInfo(I����ѹ).Text <> "" Then
                txtInfo(I����ѹ).Text = Round(Val(txtInfo(I����ѹ).Text) * 10 * 3 / 4)
            End If
            If txtInfo(I����ѹ).Text <> "" Then
                txtInfo(I����ѹ).Text = Round(Val(txtInfo(I����ѹ).Text) * 10 * 3 / 4)
            End If
        Else
            'mmHgת����Kpa �ӱ��ӱ���3�ٳ�10(Kpa������λС��)
            If txtInfo(I����ѹ).Text <> "" Then
                txtInfo(I����ѹ).Text = Format(Val(txtInfo(I����ѹ).Text) * 4 / 3 / 10, "#0.00")
            End If
            If txtInfo(I����ѹ).Text <> "" Then
                txtInfo(I����ѹ).Text = Format(Val(txtInfo(I����ѹ).Text) * 4 / 3 / 10, "#0.00")
            End If
        End If
        Call BpRange(cboBpUnit.List(cboBpUnit.ListIndex))
        cboBpUnit.Tag = cboBpUnit.List(cboBpUnit.ListIndex)
        RaiseEvent Change(7)
    End If
End Sub

Private Sub txtInfo_Change(index As Integer)
    If txtInfo(index).Text = txtInfo(index).Tag Then Exit Sub
    RaiseEvent Change(index)
End Sub

Private Sub cboBpUnit_Change()
    RaiseEvent Change(7)
End Sub

Private Sub txtInfo_Validate(index As Integer, Cancel As Boolean)
'�жϷ�Χֵ
    Dim strTmp As String
    Dim dblMin As Double
    Dim dblMax As Double
    
    If txtInfo(index).Text <> "" Then
        If Not IsNumeric(txtInfo(index).Text) Then
            MsgBox "�������ݱ��������֣�" & txtInfo(index).ToolTipText, vbInformation, "�������"
            txtInfo(index).Text = txtInfo(index).Tag
            Cancel = True
            Call zlControl.TxtSelAll(txtInfo(index))
            Exit Sub
        End If
        
        strTmp = mcol��Χ("_" & lblName(index).Tag)
        If InStr(strTmp, ";") > 0 Then
            dblMin = Val(Split(strTmp, ";")(0))
            dblMax = Val(Split(strTmp, ";")(1))
            
            If Val(txtInfo(index).Text) > dblMax Or Val(txtInfo(index).Text) < dblMin Then
                If MsgBox("��������δ��ָ����Χ�ڣ�" & txtInfo(index).ToolTipText & "�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, "�������") = vbNo Then
                    txtInfo(index).Text = txtInfo(index).Tag
                    Cancel = True
                    Call zlControl.TxtSelAll(txtInfo(index))
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

