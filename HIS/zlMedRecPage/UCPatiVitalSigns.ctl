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
      Caption         =   "���"
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   15
      Width           =   405
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
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   3615
      TabIndex        =   18
      Top             =   60
      Width           =   400
   End
   Begin VB.Label lblName 
      Caption         =   "����"
      Height          =   180
      Index           =   3
      Left            =   30
      TabIndex        =   17
      Top             =   450
      Width           =   405
   End
   Begin VB.Label lblName 
      Caption         =   "����"
      Height          =   180
      Index           =   4
      Left            =   1620
      TabIndex        =   16
      Top             =   405
      Width           =   400
   End
   Begin VB.Label lblName 
      Caption         =   "Ѫѹ"
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
      Caption         =   "��"
      Height          =   180
      Index           =   2
      Left            =   4890
      TabIndex        =   11
      Top             =   60
      Width           =   450
   End
   Begin VB.Label lblUnit 
      Caption         =   "��/��"
      Height          =   180
      Index           =   3
      Left            =   1095
      TabIndex        =   10
      Top             =   420
      Width           =   450
   End
   Begin VB.Label lblUnit 
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
'    txt��� = 0
'    txt���� = 1
'    txt���� = 2
'    txt���� = 3
'    txt���� = 4
'    txt����ѹ = 5
'    txt����ѹ = 6
'    Ѫѹ��λ = 7

Private Enum enum_txtInfo '�ı���
    txt��� = 0
    txt���� = 1
    txt���� = 2
    txt���� = 3
    txt���� = 4
    txt����ѹ = 5
    txt����ѹ = 6
End Enum

Private Enum enum_lblName '�ı�������--��Ŀ����
    lblN��� = 0
    lblN���� = 1
    lblN���� = 2
    lblN���� = 3
    lblN���� = 4
    lblNѪѹ = 5
    lblNѪѹ�ָ��� = 6
End Enum

Private Enum enum_lblUnit  '�ı���λ--��Ŀ��λ
    lblU��� = 0
    lblU���� = 1
    lblU���� = 2
    lblU���� = 3
    lblU���� = 4
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
    ControlLock = txtInfo(txt���).Locked
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
    TextBackColor = txtInfo(txt���).BackColor
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
    LblBackColor = lblName(lblN���).BackColor
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
    MaxLength = txtInfo(txt���).MaxLength
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

Public Property Get value���() As String
    value��� = txtInfo(txt���).Text
End Property

Public Property Let value���(ByVal vNewValue As String)
    txtInfo(txt���).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value���"
End Property

Public Property Get value����() As String
    value���� = txtInfo(txt����).Text
End Property

Public Property Let value����(ByVal vNewValue As String)
    txtInfo(txt����).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����"
End Property

Public Property Get value����() As String
    value���� = txtInfo(txt����).Text
End Property

Public Property Let value����(ByVal vNewValue As String)
    txtInfo(txt����).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����"
End Property

Public Property Get value����() As String
    value���� = txtInfo(txt����).Text
End Property

Public Property Let value����(ByVal vNewValue As String)
    txtInfo(txt����).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����"
End Property

Public Property Get value����() As String
    value���� = txtInfo(txt����).Text
End Property

Public Property Let value����(ByVal vNewValue As String)
    txtInfo(txt����).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����"
End Property

Public Property Get value����ѹ() As String
    value����ѹ = txtInfo(txt����ѹ).Text
End Property

Public Property Let value����ѹ(ByVal vNewValue As String)
    txtInfo(txt����ѹ).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����ѹ"
End Property

Public Property Get value����ѹ() As String
    value����ѹ = txtInfo(txt����ѹ).Text
End Property

Public Property Let value����ѹ(ByVal vNewValue As String)
    txtInfo(txt����ѹ).Text = IIf(Val(vNewValue) = 0, "", vNewValue)
    PropertyChanged "value����ѹ"
End Property

Public Property Get valueѪѹ��λ() As String
    valueѪѹ��λ = cboBpUnit.Text
End Property

Public Property Let valueѪѹ��λ(ByVal vNewValue As String)
    Call zlControl.CboLocate(cboBpUnit, vNewValue)
    PropertyChanged "valueѪѹ��λ"
End Property

Private Sub SetLine(ByVal lngStyle As Long)
'���ܣ������»���
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
            Case txt���, txt����, txt����
                strMask = "1234567890"
            Case txt����ѹ, txt����ѹ
                If cboBpUnit.Text = "mmHg" Then
                    strMask = "1234567890"
                Else
                    strMask = "1234567890."
                End If
            Case txt����, txt����
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
        lngLeft = lngLeft + txtInfo(i).Width + IIf(Me.Style = TextBox, Me.LabToTxt, 0)
        lblUnit(i).Move lngLeft, lngTop, UserControl.TextWidth("��/��"), lngHeight
        lngLeft = lngLeft + lblUnit(i).Width + Me.XDis
    Next
    
    If Me.Style = Underline Then
        For i = 0 To 5
            If InStr(lblName(i).Caption, ":") = 0 Then
                lblName(i).Caption = lblName(i).Caption & ":"
            End If
        Next
    End If
    
    lblName(lblNѪѹ�ָ���).Top = lngTop
    lblName(lblNѪѹ�ָ���).Width = UserControl.TextWidth("/")
    lblName(lblNѪѹ�ָ���).Height = lngHeight
    lblName(lblNѪѹ�ָ���).Left = txtInfo(txt����ѹ).Left + txtInfo(txt����ѹ).Width
    
    txtInfo(txt����ѹ).Move lblName(lblNѪѹ�ָ���).Left + lblName(lblNѪѹ�ָ���).Width, lngTop, UserControl.TextWidth("������"), lngHeight
    
    If Me.Style = TextBox Then
        txtInfo(txt����ѹ).Height = IIf(txtInfo(txt����ѹ).Height < 300, 300, txtInfo(txt����ѹ).Height)
        For i = 0 To 5
            txtInfo(i).Height = IIf(txtInfo(i).Height < 300, 300, txtInfo(i).Height)
            lblName(i).Top = (txtInfo(i).Height - lblName(i).Height) / 2 + txtInfo(i).Top
            lblUnit(i).Top = (txtInfo(i).Height - lblUnit(i).Height) / 2 + txtInfo(i).Top
        Next
        lblName(lblNѪѹ�ָ���).Top = (txtInfo(txt����ѹ).Height - lblName(lblNѪѹ�ָ���).Height) / 2 + txtInfo(txt����ѹ).Top
    End If
    
    cboBpUnit.Left = txtInfo(txt����ѹ).Left + txtInfo(txt����ѹ).Width
    cboBpUnit.Top = txtInfo(txt����ѹ).Top
    cboBpUnit.Width = UserControl.TextWidth("mmHg") + 400
    cboBpUnit.Height = lngHeight * 2
    
    cboBpUnit.Left = IIf(Me.Style = Underline, -30, 0)
    cboBpUnit.Top = IIf(Me.Style = Underline, -30, 0)
    
    If Me.Style = Underline Then
        fraCboBorder.Height = IIf(cboBpUnit.Height <= 300, 240, cboBpUnit.Height - 40)
        fraCboBorder.Top = txtInfo(txt����ѹ).Top
        fraCboBorder.Width = UserControl.TextWidth("mmHg") + 350
    Else
        fraCboBorder.Height = txtInfo(0).Height
        fraCboBorder.Top = txtInfo(txt����ѹ).Top
        fraCboBorder.Width = cboBpUnit.Width
    End If
    
    fraCboBorder.Left = txtInfo(txt����ѹ).Left + txtInfo(txt����ѹ).Width
    
    UserControl.Width = fraCboBorder.Left + fraCboBorder.Width
    UserControl.Height = txtInfo(txt����ѹ).Top + txtInfo(txt����ѹ).Height + 100
End Sub

Private Sub UserControl_Show()
    Call SetLine(Me.Style)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ControlLock", txtInfo(txt���).Locked, False)
    Call PropBag.WriteProperty("Enabled", Me.Enabled, True)
    Call PropBag.WriteProperty("TextBackColor", txtInfo(txt���).BackColor, &H80000000)
    Call PropBag.WriteProperty("LblBackColor", lblName(lblN���).BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", Me.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", mcolForeColor, &H80000000)
    Call PropBag.WriteProperty("BackColor", Me.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ShowMode", Me.ShowMode, enum_ShowMode.TwoRow)
    Call PropBag.WriteProperty("Style", Me.Style, enum_Style.TextBox)
    Call PropBag.WriteProperty("MaxLength", txtInfo(txt���).MaxLength, 0)
    Call PropBag.WriteProperty("XDis", mXDis, 20)
    Call PropBag.WriteProperty("YDis", mYDis, 10)
    Call PropBag.WriteProperty("LabToTxt", mLabToTxt, 10)
    
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
    Dim blnSave As Boolean
    
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
            mcol��Χ.Add strTmp, "_" & rsTmp!ID
            strTmp = Replace(strTmp, ";", " - ")
            If Val("" & rsTmp!С��) = 0 Then
                strTmp = "��ΧΪ " & strTmp & " ������"
            Else
                strTmp = "��ΧΪ " & strTmp & " ֮��������ɺ�" & rsTmp!С�� & "λС������������" & rsTmp!���� & "���ַ���"
            End If
            Select Case rsTmp!������
                Case "���"
                    lblName(lblN���).Tag = rsTmp!ID
                    lblUnit(lblU���).Caption = rsTmp!��λ
                    txtInfo(txt���).MaxLength = rsTmp!����
                    txtInfo(txt���).ToolTipText = "���" & strTmp
                Case "����"
                    lblName(lblN����).Tag = rsTmp!ID
                    lblUnit(lblU����).Caption = rsTmp!��λ
                    txtInfo(txt����).MaxLength = 5 'rsTmp!����
                    txtInfo(txt����).ToolTipText = "����" & strTmp
                Case "����"
                    lblName(lblN����).Tag = rsTmp!ID
                    lblUnit(lblU����).Caption = rsTmp!��λ
                    txtInfo(txt����).MaxLength = rsTmp!����
                    txtInfo(txt����).ToolTipText = "����" & strTmp
                Case "����"
                    lblName(lblN����).Tag = rsTmp!ID
                    lblUnit(lblU����).Caption = rsTmp!��λ
                    txtInfo(txt����).MaxLength = rsTmp!����
                    txtInfo(txt����).ToolTipText = "����" & strTmp
                Case "����"
                    lblName(lblN����).Tag = rsTmp!ID
                    lblUnit(lblU����).Caption = rsTmp!��λ
                    txtInfo(txt����).MaxLength = rsTmp!����
                    txtInfo(txt����).ToolTipText = "����" & strTmp
                Case "����ѹ"
                    lblName(lblNѪѹ).Tag = rsTmp!ID
                    Call zlControl.CboSetIndex(cboBpUnit.hwnd, IIf("" & rsTmp!��λ = "mmHg", 0, 1))
                    txtInfo(txt����ѹ).MaxLength = 5 'rsTmp!����
                    txtInfo(txt����ѹ).ToolTipText = "����ѹ" & strTmp
                Case "����ѹ"
                    lblName(lblNѪѹ�ָ���).Tag = rsTmp!ID
                    Call zlControl.CboSetIndex(cboBpUnit.hwnd, IIf("" & rsTmp!��λ = "mmHg", 0, 1))
                    txtInfo(txt����ѹ).MaxLength = 5 'rsTmp!����
                    txtInfo(txt����ѹ).ToolTipText = "����ѹ" & strTmp
            End Select
            rsTmp.MoveNext
        Next
    End If
    cboBpUnit.Tag = cboBpUnit.List(cboBpUnit.ListIndex)
    
    strSQL = "Select b.��Ŀ��λ, b.��Ŀ����, b.��¼����" & _
        " From ���˻����¼ A, ���˻������� B Where a.Id = b.��¼id And a.����id = [1] And a.��ҳid = [2]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatVitalSigns", lng����ID, lng�Һ�id)
        
    If rsTmp.RecordCount <= 0 Then
        strSQL = "Select null as ��Ŀ��λ, ��Ϣ�� As ��Ŀ����, ��Ϣֵ As ��¼���� From ������Ϣ�ӱ� Where ����id = [1] And ����id = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PatVitalSigns", lng����ID, lng�Һ�id)
        If rsTmp.RecordCount > 0 Then blnSave = True
    End If
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            Select Case rsTmp!��Ŀ����
                Case "���"
                    txtInfo(txt���).Text = rsTmp!��¼����
                    txtInfo(txt���).Tag = rsTmp!��¼����
                Case "����"
                    txtInfo(txt����).Text = rsTmp!��¼����
                    txtInfo(txt����).Tag = rsTmp!��¼����
                Case "����"
                    txtInfo(txt����).Text = rsTmp!��¼����
                    txtInfo(txt����).Tag = rsTmp!��¼����
                Case "����"
                    txtInfo(txt����).Text = rsTmp!��¼����
                    txtInfo(txt����).Tag = rsTmp!��¼����
                Case "����"
                    txtInfo(txt����).Text = rsTmp!��¼����
                    txtInfo(txt����).Tag = rsTmp!��¼����
                Case "����ѹ"
                    txtInfo(txt����ѹ).Text = rsTmp!��¼����
                    txtInfo(txt����ѹ).Tag = rsTmp!��¼����
                    If Not IsNull(rsTmp!��Ŀ��λ) Then Call zlControl.CboSetIndex(cboBpUnit.hwnd, IIf("" & rsTmp!��Ŀ��λ = "mmHg", 0, 1))
                Case "����ѹ"
                    txtInfo(txt����ѹ).Text = rsTmp!��¼����
                    txtInfo(txt����ѹ).Tag = rsTmp!��¼����
                    If Not IsNull(rsTmp!��Ŀ��λ) Then Call zlControl.CboSetIndex(cboBpUnit.hwnd, IIf("" & rsTmp!��Ŀ��λ = "mmHg", 0, 1))
                Case "Ѫѹ��λ"
                    If Not IsNull(rsTmp!��¼����) Then Call zlControl.CboSetIndex(cboBpUnit.hwnd, IIf("" & rsTmp!��¼���� = "mmHg", 0, 1))
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
'���ܣ������޸�ֵ�����棬��ָ�ԭֵ
    Dim i As Integer
    For i = 0 To 6
        txtInfo(i).Text = txtInfo(i).Tag
    Next
End Function
Public Function GetSaveSQL(ByVal lng����ID As Long, ByVal lng�Һ�id As Long) As String
'���ܣ�������������������д��SQL
    Dim strTmp As String
    Dim strSQL As String
    Dim i As Integer
    For i = 0 To 4
        strTmp = IIf(strTmp <> "", strTmp, "") & lblName(i).Tag & "|" & FormatNum(Val(txtInfo(i).Text), 2) & "|" & lblUnit(i).Caption & ","
    Next
    strTmp = strTmp & lblName(5).Tag & "|" & FormatNum(Val(txtInfo(5).Text), 2) & "|" & cboBpUnit.Text & "," & _
        lblName(6).Tag & "|" & FormatNum(Val(txtInfo(6).Text), 2) & "|" & cboBpUnit.Text
    If strTmp <> mstrValues Then
        GetSaveSQL = "Zl_������������_Update(" & lng����ID & "," & lng�Һ�id & ",'" & strTmp & "')"
    End If
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
            If txtInfo(txt����ѹ).Text <> "" Then
                txtInfo(txt����ѹ).Text = Round(Val(txtInfo(txt����ѹ).Text) * 10 * 3 / 4)
            End If
            If txtInfo(txt����ѹ).Text <> "" Then
                txtInfo(txt����ѹ).Text = Round(Val(txtInfo(txt����ѹ).Text) * 10 * 3 / 4)
            End If
        Else
            'mmHgת����Kpa �ӱ��ӱ���3�ٳ�10(Kpa������λС��)
            If txtInfo(txt����ѹ).Text <> "" Then
                txtInfo(txt����ѹ).Text = Format(Val(txtInfo(txt����ѹ).Text) * 4 / 3 / 10, "#0.00")
            End If
            If txtInfo(txt����ѹ).Text <> "" Then
                txtInfo(txt����ѹ).Text = Format(Val(txtInfo(txt����ѹ).Text) * 4 / 3 / 10, "#0.00")
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
'�жϷ�Χֵ
    Dim strTmp As String
    Dim dblMin As Double
    Dim dblMax As Double
    
    If txtInfo(Index).Text <> "" Then
        If Not IsNumeric(txtInfo(Index).Text) Then
            MsgBox "�������ݱ��������֣�" & txtInfo(Index).ToolTipText, vbInformation, "�������"
            txtInfo(Index).Text = txtInfo(Index).Tag
            Cancel = True
            Call zlControl.TxtSelAll(txtInfo(Index))
            Exit Sub
        End If
        
        strTmp = mcol��Χ("_" & lblName(Index).Tag)
        If InStr(strTmp, ";") > 0 Then
            dblMin = Val(Split(strTmp, ";")(0))
            dblMax = Val(Split(strTmp, ";")(1))
            
            If Val(txtInfo(Index).Text) > dblMax Or Val(txtInfo(Index).Text) < dblMin Then
                If MsgBox("��������δ��ָ����Χ�ڣ�" & txtInfo(Index).ToolTipText & "�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, "�������") = vbNo Then
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
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
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
