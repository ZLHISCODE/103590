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
         TabIndex        =   6
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   14
      Top             =   60
      Width           =   360
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Index           =   1
      Left            =   1350
      TabIndex        =   13
      Top             =   60
      Width           =   360
   End
   Begin VB.Label lblName 
      Caption         =   "����"
      Height          =   180
      Index           =   2
      Left            =   2970
      TabIndex        =   12
      Top             =   60
      Width           =   405
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Ѫѹ"
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
      Caption         =   "��"
      Height          =   180
      Index           =   0
      Left            =   1050
      TabIndex        =   9
      Top             =   45
      Width           =   180
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "��/��"
      Height          =   180
      Index           =   1
      Left            =   2385
      TabIndex        =   8
      Top             =   45
      Width           =   450
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      Caption         =   "��/��"
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

Public Event Change(ByVal int��� As Integer)

Private Enum E_ITEM_INDEX
    I���� = 0
    I���� = 1
    I���� = 2
    I����ѹ = 3
    I����ѹ = 4
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
Private mlng�շ�ID As Long
Private mint���� As Integer
Private mstrPreState As String 'ǰһ״̬
Private mblnNoCheck As Boolean '����ʵʱ�����ж�
Private mblnSaveNow As Boolean '�Ƿ���������

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
    ControlLock = txtInfo(I����).locked
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
    TextBackColor = txtInfo(I����).BackColor
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
    LblBackColor = lblName(I����).BackColor
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
    MaxLength = txtInfo(I����).MaxLength
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
    Call CboLocate(cboBpUnit, vNewValue)
    PropertyChanged "valueѪѹ��λ"
End Property

Private Sub SetLine(ByVal lngStyle As Long)
'���ܣ������»���
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
            Case I����, I����
                strMask = "1234567890"
            Case I����ѹ, I����ѹ
                If cboBpUnit.Text = "mmHg" Then
                    strMask = "1234567890"
                Else
                    strMask = "1234567890."
                End If
            Case I����
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
    
    Me.value���� = PropBag.ReadProperty("value����", "")
    Me.value���� = PropBag.ReadProperty("value����", "")
    Me.value���� = PropBag.ReadProperty("value����", "")
    Me.value����ѹ = PropBag.ReadProperty("value����ѹ", "")
    Me.value����ѹ = PropBag.ReadProperty("value����ѹ", "")
    Me.valueѪѹ��λ = PropBag.ReadProperty("valueѪѹ��λ", "")
    Me.Tag = PropBag.ReadProperty("Tag", "")
    
    UserControl_Resize
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
    
    For i = I���� To I����ѹ
        If i = I���� And mEnumShowMode = TwoRow Then
            lngTop = txtInfo(0).Height + Me.YDis
            lngLeft = 0
        End If
        lblName(i).Move lngLeft, lngTop, UserControl.TextWidth(IIf(Me.Style = Underline, "����:", "����")), lngHeight
        lngLeft = lngLeft + lblName(i).Width + Me.LabToTxt
        txtInfo(i).Move lngLeft, lngTop, UserControl.TextWidth("������"), lngHeight
        lngLeft = lngLeft + txtInfo(i).Width
        lblUnit(i).Move lngLeft, lngTop, IIf(mEnumShowMode = OneRow, lblUnit(i).Width, UserControl.TextWidth("��/��")), lngHeight
        If mEnumShowMode = OneRow Then
            lblUnit(i).Width = UserControl.TextWidth(lblUnit(i).Caption)
        End If
        lngLeft = lngLeft + lblUnit(i).Width + Me.XDis
    Next
    
    If mblnColon Then
        For i = I���� To I����ѹ
            If InStr(lblName(i).Caption, ":") = 0 Then
                lblName(i).Caption = lblName(i).Caption & ":"
            Else
                lblName(i).Caption = Replace(lblName(i).Caption, ":", "")
            End If
        Next
    End If
    
    lblName(I����ѹ).Top = lngTop
    lblName(I����ѹ).Width = UserControl.TextWidth("/")
    lblName(I����ѹ).Height = lngHeight
    lblName(I����ѹ).Left = txtInfo(I����ѹ).Left + txtInfo(I����ѹ).Width
    
    txtInfo(I����ѹ).Move lblName(I����ѹ).Left + lblName(I����ѹ).Width, lngTop, UserControl.TextWidth("������"), lngHeight
    
    If Me.Style = TextBox Then
        txtInfo(I����ѹ).Height = IIf(txtInfo(I����ѹ).Height < 300, 300, txtInfo(I����ѹ).Height)
        For i = I���� To I����ѹ
            txtInfo(i).Height = IIf(txtInfo(i).Height < 300, 300, txtInfo(i).Height)
            lblName(i).Top = (txtInfo(i).Height - lblName(i).Height) / 2 + txtInfo(i).Top
            lblUnit(i).Top = (txtInfo(i).Height - lblUnit(i).Height) / 2 + txtInfo(i).Top
        Next
        lblName(I����ѹ).Top = (txtInfo(I����ѹ).Height - lblName(I����ѹ).Height) / 2 + txtInfo(I����ѹ).Top
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
    UserControl.Height = txtInfo(I����ѹ).Top + txtInfo(I����ѹ).Height
End Sub

Private Sub UserControl_Show()
    Call SetLine(Me.Style)
End Sub

Private Sub UserControl_Terminate()
    mblnSaveNow = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ControlLock", txtInfo(I����).locked, False)
    Call PropBag.WriteProperty("Enabled", Me.Enabled, True)
    Call PropBag.WriteProperty("TextBackColor", txtInfo(I����).BackColor, &H80000000)
    Call PropBag.WriteProperty("LblBackColor", lblName(I����).BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", Me.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", mcolForeColor, &H80000000)
    Call PropBag.WriteProperty("BackColor", Me.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ShowMode", Me.ShowMode, enum_ShowMode.TwoRow)
    Call PropBag.WriteProperty("Style", Me.Style, enum_Style.TextBox)
    Call PropBag.WriteProperty("MaxLength", txtInfo(I����).MaxLength, 0)
    Call PropBag.WriteProperty("XDis", mXDis, 20)
    Call PropBag.WriteProperty("YDis", mYDis, 10)
    Call PropBag.WriteProperty("LabToTxt", mLabToTxt, 10)
    Call PropBag.WriteProperty("HaveColon", mblnColon, False)
    
    Call PropBag.WriteProperty("value����", Me.value����, "")
    Call PropBag.WriteProperty("value����", Me.value����, "")
    Call PropBag.WriteProperty("value����", Me.value����, "")
    Call PropBag.WriteProperty("value����ѹ", Me.value����ѹ, "")
    Call PropBag.WriteProperty("value����ѹ", Me.value����ѹ, "")
    Call PropBag.WriteProperty("valueѪѹ��λ", Me.valueѪѹ��λ, "")
    Call PropBag.WriteProperty("Tag", Me.Tag, "")
End Sub

Public Function LoadPatiVitalSigns(ByVal lng�շ�ID As Long, ByVal int���� As Integer)
'���ܣ����ؼ�¼���ݵ���Ӧ���ı����У���Ҫ�����ݱ������ı����Tagֵ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Integer
    
    mblnNoCheck = True
    
    mlng�շ�ID = lng�շ�ID
    mint���� = int����
    
    Call ClearData
    
    With cboBpUnit
        .Clear
        .AddItem "mmHg"
        .AddItem "Kpa"
    End With
    
    Set mcol��Χ = New Collection
    
    strSQL = "Select ID, ������, ����, ��λ, ��ֵ��, С�� From ����������Ŀ" & _
        " Where ����id = 7 And ������ In ('����', '����', '����ѹ', '����ѹ', '����')"
        
    On Error GoTo errH
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "PatVitalSigns")
    
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
                Case "����"
                    lblName(I����).Tag = rsTmp!id
                    lblUnit(I����).Caption = Nvl(rsTmp!��λ, "��")
                    txtInfo(I����).MaxLength = 4
                    txtInfo(I����).ToolTipText = "����" & strTmp
                Case "����"
                    lblName(I����).Tag = rsTmp!id
                    lblUnit(I����).Caption = Nvl(rsTmp!��λ, "��/��")
                    txtInfo(I����).MaxLength = 3
                    txtInfo(I����).ToolTipText = "����" & strTmp
                Case "����"
                    lblName(I����).Tag = rsTmp!id
                    lblUnit(I����).Caption = Nvl(rsTmp!��λ, "��/��")
                    txtInfo(I����).MaxLength = 3
                    txtInfo(I����).ToolTipText = "����" & strTmp
                Case "����ѹ"
                    lblName(I����ѹ).Tag = rsTmp!id
                    Call gobjControl.cbo.SetIndex(cboBpUnit.hWnd, IIf("" & rsTmp!��λ = "mmHg", 0, 1))
                    txtInfo(I����ѹ).MaxLength = 5
                    txtInfo(I����ѹ).ToolTipText = "����ѹ" & strTmp
                Case "����ѹ"
                    lblName(I����ѹ).Tag = rsTmp!id
                    Call gobjControl.cbo.SetIndex(cboBpUnit.hWnd, IIf("" & rsTmp!��λ = "mmHg", 0, 1))
                    txtInfo(I����ѹ).MaxLength = 5
                    txtInfo(I����ѹ).ToolTipText = "����ѹ" & strTmp
            End Select
            rsTmp.MoveNext
        Next
    End If
    cboBpUnit.Tag = cboBpUnit.List(cboBpUnit.ListIndex)
    
    strSQL = "select ����,����,����,����ѹ,����ѹ,Ѫѹ��λ from ѪҺִ���������� where �շ�ID=[1] and ����=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSQL, "PatVitalSigns", mlng�շ�ID, mint����)
        
    If rsTmp.RecordCount > 0 Then
        For i = 0 To rsTmp.Fields.Count - 1
            Select Case rsTmp.Fields(i).name
                Case "����"
                    txtInfo(I����).Text = "" & rsTmp.Fields(i).Value
                    txtInfo(I����).Tag = "" & rsTmp.Fields(i).Value
                Case "����"
                    txtInfo(I����).Text = "" & rsTmp.Fields(i).Value
                    txtInfo(I����).Tag = "" & rsTmp.Fields(i).Value
                Case "����"
                    txtInfo(I����).Text = "" & rsTmp.Fields(i).Value
                    txtInfo(I����).Tag = "" & rsTmp.Fields(i).Value
                Case "����ѹ"
                    txtInfo(I����ѹ).Text = "" & rsTmp.Fields(i).Value
                    txtInfo(I����ѹ).Tag = "" & rsTmp.Fields(i).Value
                Case "����ѹ"
                    txtInfo(I����ѹ).Text = "" & rsTmp.Fields(i).Value
                    txtInfo(I����ѹ).Tag = "" & rsTmp.Fields(i).Value
                Case "Ѫѹ��λ"
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
'���ܣ������޸�ֵ�����棬��ָ�ԭֵ
    Dim i As Integer
    For i = 0 To 6
        txtInfo(i).Text = txtInfo(i).Tag
    Next
End Function

Public Function GetSaveSQL(ByVal lng�շ�ID As Long, ByVal int���� As Integer) As String
'���ܣ�������������������д��SQL
    Dim strTmp As String
    Dim strSQL As String
    Dim i As Integer
    For i = I���� To I����ѹ
        If IsNumeric(txtInfo(i).Text) Then
            strTmp = gobjComlib.FormatEx(Val(txtInfo(i).Text), 2)
        Else
            strTmp = "NULL"
        End If
        strSQL = strSQL & "," & strTmp
    Next
    GetSaveSQL = "Zl_ѪҺִ����������_Update(" & lng�շ�ID & "," & int���� & strSQL & ",'" & cboBpUnit.Text & "')"
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
            For i = I����ѹ To I����ѹ
                strTmp = GetCollectContent(mcol��Χ, "_" & lblName(i).Tag)
                If InStr(strTmp, ";") > 0 Then
                    dblMin = Val(Split(strTmp, ";")(0))
                    dblMax = Val(Split(strTmp, ";")(1))
                    
                    dblMin = Round(dblMin * 10 * 3 / 4)
                    dblMax = Round(dblMax * 10 * 3 / 4)
                    txtInfo(i).ToolTipText = IIf(i = I����ѹ, "����", "����") & "ѹ��ΧΪ " & dblMin & " - " & dblMax & str��λ
                    mcol��Χ.Remove ("_" & lblName(i).Tag)
                    mcol��Χ.Add dblMin & ";" & dblMax, "_" & lblName(i).Tag
                    txtInfo(i).MaxLength = 3
                End If
            Next
        Else
            For i = I����ѹ To I����ѹ
                strTmp = GetCollectContent(mcol��Χ, "_" & lblName(i).Tag)
                If InStr(strTmp, ";") > 0 Then
                    dblMin = Val(Split(strTmp, ";")(0))
                    dblMax = Val(Split(strTmp, ";")(1))
                    
                    dblMin = Format(dblMin * 4 / 3 / 10, "#0.00")
                    dblMax = Format(dblMax * 4 / 3 / 10, "#0.00")
                    txtInfo(i).ToolTipText = IIf(i = I����ѹ, "����", "����") & "ѹ��ΧΪ " & dblMin & " - " & dblMax & str��λ
                    mcol��Χ.Remove ("_" & lblName(i).Tag)
                    mcol��Χ.Add dblMin & ";" & dblMax, "_" & lblName(i).Tag
                    txtInfo(i).MaxLength = 5
                End If
            Next
        End If
    End If
End Sub

Private Sub cboBpUnit_Click()
'Ѫѹ��λ����
    Dim strTmp As String
    
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
        RaiseEvent Change(I����ѹ + 1)
        
        If Not mblnNoCheck And mblnSaveNow Then
            strTmp = InSideSaveSQL
            If strTmp <> mstrPreState Then
                Call gobjDatabase.ExecuteProcedure(strTmp, "��������")
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
    RaiseEvent Change(I����ѹ + 1)
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
'�жϷ�Χֵ
    Dim strTmp As String
    Dim dblMin As Double
    Dim dblMax As Double
    
    If txtInfo(Index).Text <> "" Then
        If Not IsNumeric(txtInfo(Index).Text) Then
            MsgBox "�������ݱ��������֣�" & txtInfo(Index).ToolTipText, vbInformation, gstrSysName
            txtInfo(Index).Text = txtInfo(Index).Tag
            Cancel = True
            Call gobjControl.TxtSelAll(txtInfo(Index))
            Exit Sub
        End If
        
        strTmp = GetCollectContent(mcol��Χ, "_" & lblName(Index).Tag)
        If InStr(strTmp, ";") > 0 Then
            dblMin = Val(Split(strTmp, ";")(0))
            dblMax = Val(Split(strTmp, ";")(1))
            
            If Val(txtInfo(Index).Text) > dblMax Or Val(txtInfo(Index).Text) < dblMin Then
                MsgBox "��������δ��ָ����Χ�ڣ�" & txtInfo(Index).ToolTipText, vbInformation, gstrSysName
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
            Call gobjDatabase.ExecuteProcedure(strTmp, "��������")
            mstrPreState = strTmp
        End If
    End If
End Sub

Public Function ClearTxtToolTipText()
'���ܣ���������ı������ʾ
    Dim i As Integer
    
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).ToolTipText = ""
    Next
End Function

Public Sub TxtAlignment(ByVal intType As Integer)
'���ܣ������ı���Ķ��뷽ʽ
'intType 0-����룬1���Ҷ��룬2������
    Dim i As Long
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).Alignment = intType
    Next
End Sub

Public Sub SetUseType(ByVal blnSaveNow As Boolean, Optional ByVal strTag As String)
    mblnSaveNow = blnSaveNow
End Sub

Private Function InSideSaveSQL() As String
'���ܣ�������������������д��SQL
    Dim strTmp As String
    Dim strSQL As String
    Dim i As Integer
    If Not mblnSaveNow Then Exit Function
    For i = I���� To I����ѹ
        If IsNumeric(txtInfo(i).Text) Then
            strTmp = gobjComlib.FormatEx(Val(txtInfo(i).Text), 2)
        Else
            strTmp = "NULL"
        End If
        strSQL = strSQL & "," & strTmp
    Next
    InSideSaveSQL = "Zl_ѪҺִ����������_Update(" & mlng�շ�ID & "," & mint���� & strSQL & ",'" & cboBpUnit.Text & "')"
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

