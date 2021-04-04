VERSION 5.00
Begin VB.UserControl PatiAddress 
   BackColor       =   &H80000005&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   ScaleHeight     =   360
   ScaleWidth      =   4920
   Begin VB.TextBox txtInfo 
      ForeColor       =   &H80000000&
      Height          =   300
      Index           =   4
      Left            =   3480
      TabIndex        =   4
      Text            =   "��ϸ��ַ"
      Top             =   30
      Width           =   1417
   End
   Begin VB.TextBox txtInfo 
      ForeColor       =   &H80000000&
      Height          =   300
      Index           =   3
      Left            =   2625
      TabIndex        =   3
      Text            =   "��(��)"
      Top             =   30
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox txtInfo 
      ForeColor       =   &H80000000&
      Height          =   300
      Index           =   2
      Left            =   1740
      TabIndex        =   2
      Text            =   "��(��)"
      Top             =   30
      Width           =   945
   End
   Begin VB.TextBox txtInfo 
      ForeColor       =   &H80000000&
      Height          =   300
      Index           =   1
      Left            =   850
      TabIndex        =   1
      Text            =   "��"
      Top             =   30
      Width           =   945
   End
   Begin VB.TextBox txtInfo 
      ForeColor       =   &H80000000&
      Height          =   300
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Text            =   "ʡ(��,��)"
      Top             =   30
      Width           =   945
   End
   Begin VB.Menu mnuPopuMenu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuMenuCopyAll 
         Caption         =   "����������ַ"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuPopuMenuCopy 
         Caption         =   "����"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPopuMenuPasteAll 
         Caption         =   "ճ��������ַ"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuPopuMenuPaste 
         Caption         =   "ճ��"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuPopuMenuDelete 
         Caption         =   "��յ�ַ"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "PatiAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Enum enum_txtInfo
    txtʡ = 0
    txt�� = 1
    txt���� = 2
    txt���� = 3
    txt��ϸ��ַ = 4
End Enum

Public Enum enum_Items
    One = 1
    Two = 2
    Three = 3
    Four = 4
    Five = 5
End Enum

Public Enum enum_Style
    TextBox = 0
    Underline = 1
End Enum

Private Type ItemInfo
    strInfo As String 'ƥ��ĵ�ַ����
    strCode As String 'ƥ��ĵ�ַ����
    strNullInfo As String 'û������ʱĬ����ʵ
    strStName  As String '��׼����
    blnƥ�� As Boolean '�Ƿ񾭹�ƥ�����
    bln���� As Boolean '�Ƿ��������ַ
    bln����ʾ As Boolean '�Ƿ������ⲻ��ʾ���ݵĵ�ַ
    bln��Ч As Boolean '�Ƿ�δʹ��
    bln���� As Boolean '�Ƿ���������
End Type

'���Ա���
Private mstrTag As String
Private mblnShowTown As Boolean 'ShowTown����
Private mblnLocked As Boolean 'ControlLock����
Private mcolForeColor As OLE_COLOR
Private mEnumStyle As enum_Style
Private mEnumItemCount As enum_Items
Private mtxtBackColor As OLE_COLOR

'�ڲ�����
Private marrItems(4) As ItemInfo
Private mstrLike As String
Private mblnLike As Boolean
Private mblnFocus As Boolean
Private mblnCancel As Boolean
Private mblnResize As Boolean '��ֹѭ������Resize
Private mblnSetItems As Boolean '��ֹѭ������TxtInfo_change
Private mblnChange As Boolean
Private mstrOldAddress As String
Private mblnChangeOld As Boolean '�Ƿ��޸��ϵ�ַ
Private mblnEdit As Boolean    '�Ƿ�༭�ɹ�
Private mblnLineFeed As Boolean '��ϸ��ַ�Ƿ�����ʾ

Public Event Change()
Public Event SetEdit(blnEdit As Boolean)
Public Event SetInput(ByVal intLevel As Integer, rsReturn As ADODB.Recordset)

'==============================================================
'===�Զ���ؼ�����
'==============================================================
'hwnd:���ھ��
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property
'Style:�ؼ������ʽ
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
    PropertyChanged "Style"
End Property
'Enabled:�ؼ�����״̬
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    Dim i As Long
    UserControl.Enabled = NewValue
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).Enabled = NewValue
        txtInfo(i).BackColor = IIf(NewValue, vbWindowBackground, vbButtonFace)
    Next
    If Not NewValue Then Me.ControlLock = True
    PropertyChanged "Enabled"
End Property

'ControlLock:�ؼ���Lock״̬
Public Property Get ControlLock() As Boolean
    ControlLock = mblnLocked
End Property

Public Property Let ControlLock(ByVal NewValue As Boolean)
    Dim i As Long
    mblnLocked = NewValue
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).Locked = NewValue
        txtInfo(i).TabStop = Not NewValue
        If txtInfo(i).Enabled Then txtInfo(i).BackColor = IIf(NewValue, vbButtonFace, vbWindowBackground)
    Next
    PropertyChanged "ControlLock"
End Property

Public Property Get LineFeed() As Boolean
    LineFeed = mblnLineFeed
End Property

Public Property Let LineFeed(ByVal NewValue As Boolean)
    Dim i As Long
    mblnLineFeed = NewValue
    If Items() = Four Or Items() = Five Then
        Call UserControl_Resize
    End If
    PropertyChanged "LineFeed"
End Property
'Font:�ؼ�����
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Dim i As Long
    For i = 0 To txtInfo.Count - 1
        Set txtInfo(i).Font = New_Font
    Next
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Call UserControl_Resize
End Property
'Tag:�洢�ؼ���صĶ�������
Public Property Get Tag() As String
    Tag = mstrTag
End Property

Public Property Let Tag(ByVal vNewValue As String)
    mstrTag = vNewValue
    PropertyChanged "Tag"
End Property
'ForeColor:�ؼ���������ɫ
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
'TextBackColor:�����ı�����ɫ
Public Property Get TextBackColor() As OLE_COLOR
    TextBackColor = mtxtBackColor
End Property

Public Property Let TextBackColor(ByVal New_BackColor As OLE_COLOR)
    Dim i As Long
    mtxtBackColor = New_BackColor
    For i = 0 To txtInfo.Count - 1
       txtInfo(i).BackColor = New_BackColor
    Next
    PropertyChanged "TextBackColor"
End Property
'BackColor:�ؼ��ı�����ɫ
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property
'MaxLength:������Ԫ�������������󳤶�
Public Property Get MaxLength() As Long
    MaxLength = txtInfo(txtʡ).MaxLength
End Property

Public Property Let MaxLength(ByVal vNewValue As Long)
    Dim i As Long
    For i = 0 To txtInfo.Count - 1
        txtInfo(i).MaxLength = vNewValue
    Next
    PropertyChanged "MaxLength"
End Property
'Items:�ؼ�չʾ����Ŀ����
Public Property Get Items() As enum_Items
    Items = IIf(mEnumItemCount = 0, enum_Items.Four, mEnumItemCount)
End Property

Public Property Let Items(ByVal vNewValue As enum_Items)
    Dim i As Long, lngCount As Long
    
    For i = 0 To txt��ϸ��ַ
        marrItems(i).bln��Ч = Not (i < vNewValue)
    Next
    If vNewValue = Four Then
        marrItems(txt��ϸ��ַ).bln��Ч = False
        marrItems(txt����).bln��Ч = True
    End If
    mEnumItemCount = vNewValue
    PropertyChanged "Items"
    If vNewValue = Five And Not mblnShowTown Or vNewValue <> Four And mblnShowTown Then
        mblnShowTown = (vNewValue = Five)
        PropertyChanged "ShowTown"
    End If
    Call UserControl_Resize
    mblnChange = False
End Property
'ShowTown:�Ƿ�չʾ���򼶣�Items=4ʱ������
Public Property Get ShowTown() As Boolean
    ShowTown = mblnShowTown
End Property

Public Property Let ShowTown(ByVal vNewValue As Boolean)
    '5�����ļ�����ʾ������5����ַ�������ļ��ĵ�ַ����������ԶΪFalse
    Dim i As Integer

    If mEnumItemCount = Four And vNewValue Or mEnumItemCount = Five And Not vNewValue Then
        mEnumItemCount = IIf(vNewValue, Five, Four)
    ElseIf vNewValue <> (mEnumItemCount = Five) Then '�弶��ַ
        mblnShowTown = mEnumItemCount = Five
    Else
        mblnShowTown = vNewValue
    End If
    For i = 0 To txt��ϸ��ַ
        marrItems(i).bln��Ч = Not (i < Me.Items)
    Next
    If Me.Items = Four Then
        marrItems(txt��ϸ��ַ).bln��Ч = False
        marrItems(txt����).bln��Ч = True
    End If
    PropertyChanged "ShowTown"
    PropertyChanged "Items"
    Call UserControl_Resize
End Property
'value:�ؼ���ֵ
Public Property Get value() As String
    value = Me.valueʡ & Me.value�� & Me.value���� & Me.value���� & Me.value��ϸ��ַ
End Property

Public Property Let value(ByVal vNewValue As String)
    Call LoadAllAdress(vNewValue, Me.Items)
    PropertyChanged "value"
    Call UserControl_Resize
End Property
'valueʡ:ʡ��(��һ��)��ַ��ֵ
Public Property Get valueʡ() As String
    valueʡ = marrItems(txtʡ).strInfo
End Property

'value��:�м�(�ڶ���)��ַ��ֵ
Public Property Get value��() As String
    value�� = IIf(marrItems(txt��).bln����ʾ Or marrItems(txt��).bln��Ч, "", marrItems(txt��).strInfo)
End Property

'value����:���ؼ�(������)��ַ��ֵ
Public Property Get value����() As String
    value���� = IIf(marrItems(txt����).bln����ʾ Or marrItems(txt����).bln��Ч, "", marrItems(txt����).strInfo)
End Property

'value����:����(���ļ�)��ַ��ֵ
Public Property Get value����() As String
    value���� = IIf(marrItems(txt����).bln����ʾ Or marrItems(txt����).bln��Ч, "", marrItems(txt����).strInfo)
End Property

'value��ϸ��ַ:���(���弶)��ַ��ֵ
Public Property Get value��ϸ��ַ() As String
    value��ϸ��ַ = IIf(marrItems(txt��ϸ��ַ).bln����ʾ Or marrItems(txt��ϸ��ַ).bln��Ч, "", marrItems(txt��ϸ��ַ).strInfo)
End Property

'Code:��׼����ַ��Ԫ������Сһ����ַ��Ӧ�ı���
Public Property Get Code() As String
    Dim i As Integer, strTmp As String
    For i = txt��ϸ��ַ To 0 Step -1
        strTmp = marrItems(i).strCode
        If strTmp <> "" Then Exit For
    Next
    Code = strTmp
End Property
'AllCodes:���е�ַ��Ԫ���Ӧ���룬�Զ��ŷָ�
Public Property Get AllCodes() As String
    AllCodes = marrItems(txtʡ).strCode & "," & marrItems(txt��).strCode & "," & marrItems(txt����).strCode & "," & marrItems(txt����).strCode & "," & marrItems(txt��ϸ��ַ).strCode
End Property

'==============================================================
'===�ؼ�����
'==============================================================
Public Sub LoadAllAdress(ByVal strAdress As String, Optional ByVal intType As Integer)
    Call StructAdress(strAdress, intType)
End Sub

Public Sub LoadStructAdress(ByVal strʡ As String, ByVal str�� As String, ByVal str���� As String, ByVal str���� As String, ByVal str��ϸ��ַ As String, Optional ByVal intType As Integer)
   Call StructAdress(strʡ & "," & str�� & "," & str���� & "," & str���� & "," & str��ϸ��ַ, intType)
End Sub

Public Function CheckNullValue(Optional ByVal blnNotCheck��ϸ��ַ As Boolean = True, Optional ByVal blnOnlyChangeCheck As Boolean, Optional ByVal blnMustInput As Boolean) As String
'���ܣ���������ʱ���п�ֵ��飬��֤����˳�����롣
'������blnOnlyChangeCheck=ֻ�б仯�ż��,Ϊ�գ��ұ�������ʱ���ò�������ʶ���ϵ�ַ�����⣬������
'          blnMustInput=�Ƿ��������
'          blnNotCheck��ϸ��ַ=��ϸ��ַû���Ƿ񲻼��,
'˵����
    Dim i As Long, blnNull As Boolean
    Dim blnCheck As Boolean
    
    If Me.value = "" And blnMustInput Then
        blnCheck = True
    ElseIf Me.value <> "" Then
        If blnOnlyChangeCheck Then
            blnCheck = mstrOldAddress <> Me.value
        Else
            blnCheck = True
        End If
    End If
    
    If blnCheck Then
        For i = 0 To txt��ϸ��ַ
            If Not marrItems(i).bln��Ч And txtInfo(i).Visible Then
                If marrItems(i).strInfo = "" Then
                    If Not (i = txt�� And InStr(marrItems(0).strInfo, "��") > 0) Then
                        If i = txt��ϸ��ַ And Not blnNotCheck��ϸ��ַ Or i <> txt��ϸ��ַ And i <> txt���� Then
                            CheckNullValue = marrItems(i).strNullInfo
                            Exit For
                        End If
                    End If
                End If
            End If
        Next
    End If
End Function

Public Function CheckDefrentValue(ByVal NowAdress As String, Optional ByVal PreAdress As String = "") As Boolean
'���ܣ����֤��ַ��¼��ĵ�ַ����У��
'������NowAdress=���ݿ��ȡ�����Ļ��ڵ�ַ��Ϣ
'          PreAdress=������¼��ĵ�ַ��Ϣ
    Dim i As Integer
    Dim strPatiAdress As String
    If Trim(NowAdress) = "" Then
        CheckDefrentValue = True
    ElseIf Trim(NowAdress) <> "" Then
        If Me.Items >= Four Then
            strPatiAdress = Trim(NowAdress)
            Me.value = PreAdress
            If Me.value = strPatiAdress Then
                CheckDefrentValue = True
            Else
                CheckDefrentValue = False
            End If
            Me.value = strPatiAdress
        End If
    End If
    Exit Function
End Function

'==============================================================
'===�Զ���ؼ��¼�
'==============================================================
Private Sub UserControl_GotFocus()
    Set gobjPati = Me
    If txtInfo(txtʡ).Enabled Then Call txtInfo(txtʡ).SetFocus
End Sub

Private Sub UserControl_Initialize()
    Dim i As Integer
    marrItems(txtʡ).strNullInfo = "ʡ(��,��)"
    marrItems(txt��).strNullInfo = "��"
    marrItems(txt����).strNullInfo = "��(��)"
    marrItems(txt����).strNullInfo = "��(��)"
    marrItems(txt��ϸ��ַ).strNullInfo = "��ϸ��ַ"
End Sub

Private Sub UserControl_InitProperties()
    If mEnumItemCount = 0 Then
        mEnumItemCount = Four
        mblnShowTown = False
        marrItems(txt����).bln��Ч = True
    End If
    mEnumStyle = enum_Style.TextBox
    mtxtBackColor = &H80000005
    mcolForeColor = &H80000000
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyE And Shift = vbCtrlMask Then
        If Me.value <> "" Then
            Call mnuPopuMenuDelete_Click
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If Me.value <> "" Then
            Call mnuPopuMenuCopyAll_Click
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyW And Shift = vbCtrlMask Then
        If Clipboard.GetText <> "" Then
            Call mnuPopuMenuPasteAll_Click
            KeyCode = 0
        End If
    ElseIf KeyCode = vbKeyV And Shift = vbCtrlMask Then
        If Clipboard.GetText Like "ZLSOFT:*" Then
            Call mnuPopuMenuPasteAll_Click
        End If
    End If
End Sub

Private Sub UserControl_Paint()
    Call SetLine(Me.Style)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mblnResize = True
    Me.Style = PropBag.ReadProperty("Style", enum_Style.TextBox)
    Me.Items = PropBag.ReadProperty("Items", enum_Items.Four)
    Me.ShowTown = PropBag.ReadProperty("ShowTown", Me.ShowTown)  '��ΪItems�������������أ����Ĭ��ֵΪMe.ShowTown
    Me.ControlLock = PropBag.ReadProperty("ControlLock", False)
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    Me.TextBackColor = PropBag.ReadProperty("TextBackColor", &H80000005)
    Set Me.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Me.ForeColor = PropBag.ReadProperty("ForeColor", &H80000000)
    Me.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Me.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Me.value = PropBag.ReadProperty("value", "")
    Me.Tag = PropBag.ReadProperty("Tag", "")
    mblnResize = False: mblnChangeOld = True
    Me.LineFeed = PropBag.ReadProperty("LineFeed", False)
    Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
'����: ���ÿؼ���Сλ��
    On Error Resume Next
    Dim i As Long, intPreVisual As Integer
    Dim lngHeight As Long, lngPerWidth As Long, lngMinWidth As Long, lngMinHeight As Long
    Dim arrWidthShare As Variant, lngTotal As Double
    Dim lngCount As Long, lngLastItem As Long
    Dim lngDisH As Long, lngDisV As Long

    If mblnResize Then Exit Sub
    mblnResize = True
    For i = 0 To txt��ϸ��ַ
        txtInfo(i).Visible = Not marrItems(i).bln��Ч
    Next
    If Me.Items = Two Then
        txtInfo(txt��).Visible = marrItems(txt����).bln��Ч
    End If
    lngMinWidth = UserControl.TextWidth("ʡ(��,��)") + 60
    '�����ı����ȱ��������ĸ���ĿΪ�����׼���弶��ַʱͨ�����ļ����
    arrWidthShare = Array(1, 1, 1, 2.5)
    lngCount = Me.Items
    lngLastItem = lngCount - 1
    If Me.Items = Four Then lngLastItem = 4
    If Me.Items = Five Then lngCount = 4
    For i = 0 To lngCount - 1
        If mblnLineFeed And Me.Items <> Five Then
            If i <> 3 And i <> 4 Then
                lngTotal = lngTotal + arrWidthShare(i)
            End If
        Else
            lngTotal = lngTotal + arrWidthShare(i)
        End If
    Next
    lngDisH = 0: lngDisV = 0
    lngPerWidth = (UserControl.Width - (lngCount + 1) * lngDisH) / lngTotal
    If lngPerWidth < lngMinWidth Then lngPerWidth = lngMinWidth
    lngMinHeight = UserControl.TextHeight("��")
    If UserControl.Height >= txtInfo(txtʡ).Height And mblnLineFeed And lngCount = 4 Then
        lngHeight = UserControl.Height / 2
    Else
        lngHeight = UserControl.Height - IIf(Me.Style = Underline, 30, 0)
    End If
    If lngHeight < lngMinHeight Then lngHeight = lngMinHeight
    '�ؼ�λ�ð���
    For i = 0 To txtInfo.Count - 1
        If i = 0 Then
            txtInfo(i).Move lngDisH, lngDisV, lngPerWidth * arrWidthShare(i), lngHeight
        ElseIf i < lngCount Then
            If mblnLineFeed And Me.Items <> Five Then
                If i <> 3 And i <> 4 Then
                    txtInfo(i).Move txtInfo(i - 1).Left + txtInfo(i - 1).Width + lngDisH, lngDisV, lngPerWidth * arrWidthShare(i), lngHeight
                End If
            Else
                txtInfo(i).Move txtInfo(i - 1).Left + txtInfo(i - 1).Width + lngDisH, lngDisV, lngPerWidth * arrWidthShare(i), lngHeight
            End If
        Else '��չʾ�Ŀؼ�Ҳ��Ҫ����λ�ã���ֹ���߳�����
            If Me.Items = Four Or Me.Items = Five Then
                If mblnLineFeed Then
                    If i = txt��ϸ��ַ Then
                        txtInfo(i).Move lngDisH, lngDisV + lngHeight, lngPerWidth * 2.5, lngHeight
                    Else
                        txtInfo(i).Move lngDisH, lngDisV + lngHeight, lngPerWidth * 1, lngHeight
                    End If
                Else
                    txtInfo(i).Move txtInfo(i - 1).Left + txtInfo(i - 1).Width + lngDisH, lngDisV, lngPerWidth * 1, lngHeight
                End If
            End If
        End If
    Next
    If Me.Items = Four Then
        If mblnLineFeed Then
            txtInfo(txt��ϸ��ַ).Move lngDisH, lngDisV + lngHeight, txtInfo(txtʡ).Width + txtInfo(txt��).Width + txtInfo(txt����).Width, lngHeight
        Else
            txtInfo(txt��ϸ��ַ).Move txtInfo(txt����).Left + txtInfo(txt����).Width + lngDisH, lngDisV, txtInfo(txt����).Width, lngHeight
        End If
    ElseIf Me.Items = Five Then
        '��������ϸ��ַ����Ϊ1:1.5
        lngPerWidth = (txtInfo(txt����).Width - lngDisH) / (1 + 1.5)
        If mblnLineFeed Then
            txtInfo(txt����).Width = lngPerWidth * 2.5
        Else
            txtInfo(txt����).Width = lngPerWidth * 1
        End If
        If mblnLineFeed Then
            txtInfo(txt��ϸ��ַ).Move lngDisH, lngDisV + lngHeight, txtInfo(txtʡ).Width + txtInfo(txt��).Width + txtInfo(txt����).Width + txtInfo(txt����).Width, lngHeight
        Else
            txtInfo(txt��ϸ��ַ).Move txtInfo(txt����).Left + txtInfo(txt����).Width + lngDisH, lngDisV, lngPerWidth * 1.5, lngHeight
        End If
    ElseIf Me.Items = Two Then
        If txtInfo(txt��).Visible Then
            txtInfo(txt����).Move txtInfo(txt��).Left + txtInfo(txt��).Width, 0, txtInfo(txt��).Width, txtInfo(txt��).Height
            txtInfo(txt��).ZOrder
        Else
            txtInfo(txt����).Move txtInfo(txt��).Left, 0, txtInfo(txt��).Width, txtInfo(txt��).Height
            txtInfo(txt����).ZOrder
        End If
    End If
    If mblnLineFeed Then
        If Me.Items = Four Or Me.Items = Five Then
            UserControl.Height = txtInfo(txtʡ).Top + txtInfo(txtʡ).Height * 2 + IIf(Me.Style = Underline, 30, 0)
            If Me.Items = Four Then
                UserControl.Width = txtInfo(txtʡ).Width + txtInfo(txt��).Width + txtInfo(txt����).Width
            Else
                UserControl.Width = txtInfo(txtʡ).Width + txtInfo(txt��).Width + txtInfo(txt����).Width + txtInfo(txt����).Width
            End If
        Else
            UserControl.Width = txtInfo(lngLastItem).Left + txtInfo(lngLastItem).Width + lngDisH
            UserControl.Height = txtInfo(txtʡ).Top + txtInfo(txtʡ).Height + IIf(Me.Style = Underline, 30, 0)
        End If
    Else
        UserControl.Height = txtInfo(txtʡ).Top + txtInfo(txtʡ).Height + IIf(Me.Style = Underline, 30, 0)
        UserControl.Width = txtInfo(lngLastItem).Left + txtInfo(lngLastItem).Width + lngDisH
    End If
    UserControl.Refresh
    Call SetLine(Me.Style)
    mblnResize = False
End Sub

Private Sub UserControl_Show()
    Call SetLine(Me.Style)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ControlLock", mblnLocked, False)
    Call PropBag.WriteProperty("Enabled", Me.Enabled, True)
    Call PropBag.WriteProperty("TextBackColor", mtxtBackColor, &H80000005)
    Call PropBag.WriteProperty("Font", Me.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", mcolForeColor, &H80000000)
    Call PropBag.WriteProperty("BackColor", Me.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Items", Me.Items, enum_Items.Four)
    Call PropBag.WriteProperty("ShowTown", Me.ShowTown, mblnShowTown)
    Call PropBag.WriteProperty("Style", Me.Style, enum_Style.TextBox)
    Call PropBag.WriteProperty("MaxLength", txtInfo(txtʡ).MaxLength, 0)
    Call PropBag.WriteProperty("value", Me.value, "")
    Call PropBag.WriteProperty("Tag", Me.Tag, "")
    Call PropBag.WriteProperty("LineFeed", mblnLineFeed, False)
End Sub
'==============================================================
'===�Զ���ؼ��ڲ��ռ��¼�
'==============================================================
Private Sub mnuPopuMenuCopyAll_Click()
    Dim i As Long, strAdress As String
    Dim strTmp As String
    
    Clipboard.Clear
    For i = 0 To txt��ϸ��ַ
        strTmp = marrItems(i).strInfo & "," & marrItems(i).strCode & "," & IIf(marrItems(i).bln����, 1, 0) & "," & _
                        IIf(marrItems(i).bln����ʾ, 1, 0) & ",," & IIf(marrItems(i).blnƥ��, 1, 0)
        strAdress = strAdress & IIf(strAdress = "", "", "|") & strTmp
    Next
    strAdress = "ZLSOFT:" & strAdress
    Clipboard.SetText strAdress
End Sub

Private Sub mnuPopuMenuCopy_Click()
    Clipboard.Clear
    If Not UserControl.ActiveControl Is Nothing Then
         Clipboard.SetText UserControl.ActiveControl.SelText
    End If
End Sub

Private Sub mnuPopuMenuDelete_Click()
    Dim i As Long, intCur As Integer
    If Not (Me.Enabled And Not Me.ControlLock) Then Exit Sub
    intCur = -1
    If Not UserControl.ActiveControl Is Nothing Then
        intCur = UserControl.ActiveControl.Index
    End If
    For i = 0 To txt��ϸ��ַ
        txtInfo(i).Text = ""
        Call FillItems(i, , i = intCur)
    Next
End Sub

Private Sub mnuPopuMenuPasteAll_Click()
    Dim i As Long, intCur As Integer
    Dim strTmp As String
    
    
    If Not (Me.Enabled And Not Me.ControlLock) Then Exit Sub
    intCur = -1
    If Not UserControl.ActiveControl Is Nothing Then
        intCur = UserControl.ActiveControl.Index
    End If
    strTmp = Clipboard.GetText
    If zlCommFun.ActualLen(strTmp) > 500 Then
        If Clipboard.GetText Like "ZLSOFT:*" Then
        Else
            strTmp = SubB(strTmp, 1, 500)
        End If
    End If
    mblnChangeOld = False
    Call StructAdress(strTmp)
    For i = 0 To txt��ϸ��ַ
        Call FillItems(i, , i = intCur)
    Next
    mblnChangeOld = True
End Sub

Private Sub mnuPopuMenuPaste_Click()
    If Not (Me.Enabled And Not Me.ControlLock) Then Exit Sub
    If Not UserControl.ActiveControl Is Nothing Then
        If Clipboard.GetText Like "ZLSOFT:*" Then
            Call mnuPopuMenuPasteAll_Click
        Else
            Call SendMessage(UserControl.ActiveControl.hWnd, WM_PASTE, 0, 0)
        End If
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If mblnSetItems Then Exit Sub
    marrItems(Index).strInfo = txtInfo(Index).Text
    Call ClearItems(Index)
    RaiseEvent Change
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Set gobjPati = Me
    Call FillItems(Index)
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If KeyCode = 13 Or ((KeyCode = vbKeySpace Or Chr(KeyCode) = "*") And txtInfo(Index).Text = "") Then
        '������������
       
       If (txtInfo(Index).Tag = marrItems(Index).strCode And marrItems(Index).blnƥ�� Or txtInfo(Index).Text = "") And KeyCode = 13 Then
            KeyCode = 0
            If txtInfo(Index).Text = "" And Not marrItems(Index).blnƥ�� Then
                Call ClearItems(Index)
            End If
            Call LocateItem(Index, 1, marrItems(Index).blnƥ��)
            Exit Sub
        Else
            KeyCode = 0
            Call SetInput(Index)
        End If
    ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyDown Or KeyCode = vbKeyTab Then
        KeyCode = 0
        Call LocateItem(Index, 1)
    ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
        KeyCode = 0
        Call LocateItem(Index, -1)
    ElseIf KeyCode = vbKeyV And Shift = vbCtrlMask Then
        gblnCanPaste = txtInfo(Index).Enabled And Not txtInfo(Index).Locked
        If gblnCanPaste And Clipboard.GetText Like "ZLSOFT:*" Then gblnCanPaste = False
        If glngTXTProc = 0 Then
            glngTXTProc = GetWindowLong(txtInfo(Index).hWnd, GWL_WNDPROC)
            Call SetWindowLong(txtInfo(Index).hWnd, GWL_WNDPROC, AddressOf WndMessagePaste)
        End If
    End If
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        KeyAscii = 0
        txtInfo(Index).Text = ""
    End If
End Sub

Private Sub txtInfo_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyV And Shift = vbCtrlMask Then
        If glngTXTProc <> 0 Then
            Call SetWindowLong(txtInfo(Index).hWnd, GWL_WNDPROC, glngTXTProc)
            glngTXTProc = 0
        End If
    End If
End Sub

Private Sub txtInfo_LostFocus(Index As Integer)
    txtInfo(Index).SelStart = 0
    txtInfo(Index).SelLength = 0
    Call FillItems(Index, , False)
End Sub

Private Sub txtInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call TxtMouseDown(txtInfo(Index), Button, Shift, X, Y)
End Sub

Private Sub txtInfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtInfo(Index).ToolTipText = txtInfo(Index).Text
End Sub

Private Sub txtInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call TxtMouseUp(txtInfo(Index), Button, Shift, X, Y)
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    If Not marrItems(Index).blnƥ�� And txtInfo(Index).Text <> marrItems(Index).strNullInfo And txtInfo(Index).Text <> "" Then
        Cancel = SetInput(Index)
    End If
End Sub
'==============================================================
'===�ڲ�����
'==============================================================
Private Function SetInput(ByVal intIndex As Integer, Optional ByVal strInputCode As String, Optional ByRef strPreCode As String, Optional ByVal strName As String, Optional ByVal blnClare As Boolean = True) As Boolean
'���ܣ������������������ı�������
' ������intIndex:���д���Ŀؼ�
'          strInputCode:="",�Ե�ǰ��Ԫ���������ƥ�䣬<>"" �Ե�ǰ��Ԫ����о�ȷ���ң��������strPreCode��strName��Ϊ�գ��򲻻���в��ң�ֱ�Ӽ��أ�
'          strPreCode=�ϼ�����
'          strName=��ǰ��Ԫ��ƥ�䵽������
'          blnClare=�Ƿ�����䶯��Ŀ
' ���أ�
'         strPreCode=�ϼ�����
'         �Ƿ��ֹ����ƶ�
    Dim intPreIndex As Integer
    Dim strTmpSQL As String, strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strInput As String, strCode As String, i As Integer
    Dim vPoint As POINTAPI, blnCancel As Boolean
    Dim intLevel As Integer
    
    If strInputCode = "" Then
        intPreIndex = GetSetableItem(intIndex, -1)
        strInput = Trim(txtInfo(intIndex).Text)
        If intPreIndex >= 0 Then strCode = marrItems(intPreIndex).strCode
        intLevel = intIndex
        If strCode = "" And strInput = "" And (intIndex = txt��ϸ��ַ Or intIndex = txt����) Then Exit Function
        '��������ʱ������ϸ��ַ����
        If intIndex = txt��ϸ��ַ And marrItems(txt����).bln��Ч Then
            intLevel = txt����
            If intPreIndex + 1 = intLevel Then
                strTmpSQL = IIf(strCode <> "", " And B.�ϼ�����=[4]", "")
            Else
                strTmpSQL = IIf(strCode <> "", " And B.�ϼ����� In(Select D.���� From ���� D Where D.�ϼ�����=[4])", "")
            End If
            
            If strInput <> "" Then
                If zlCommFun.IsCharChinese(strInput) Then
                    strTmpSQL = strTmpSQL & " And Nvl(a.����,b.����) Like [1] "
                ElseIf IsNumeric(strInput) Then
                    strTmpSQL = strTmpSQL & " And Nvl(a.����,b.����) Like [1]  "
                Else
                    strTmpSQL = strTmpSQL & " And Nvl(a.����,b.����) Like [1] "
                End If
            End If
            strSQL = "Select Rownum as ID,Nvl(a.����,b.����) ����, b.���� || a.���� ����, Nvl(a.����,b.����) ����, b.�ϼ�����, a.�Ƿ�����, a.�Ƿ���ʾ" & vbNewLine & _
                            "From ���� a, ���� b" & vbNewLine & _
                            "Where a.�ϼ�����(+) = b.���� And NVL(B.����,0)=[3] " & strTmpSQL & " Order by Nvl(a.����,b.����)"
        Else
            If intPreIndex + 1 = intLevel Then
                strTmpSQL = IIf(strCode <> "", " And A.�ϼ�����=[4]", "")
            Else
                strTmpSQL = IIf(strCode <> "", " And A.�ϼ����� In(Select ���� From ����  Where �ϼ�����=[4])", "")
            End If
            If strInput <> "" Then
                If zlCommFun.IsCharChinese(strInput) Then
                    strTmpSQL = strTmpSQL & " And A.���� Like [1] "
                ElseIf IsNumeric(strInput) Then
                    strTmpSQL = strTmpSQL & " And A.���� Like [1]  "
                Else
                    strTmpSQL = strTmpSQL & " And A.���� Like [1] "
                End If
            End If
            
            strSQL = "Select Rownum as ID,����,����,����,�ϼ�����,�Ƿ�����,�Ƿ���ʾ,�ʱ�  From ���� A " & _
                            "Where NVL(A.����,0)=[3]" & strTmpSQL & " Order by ����"
        End If
        If mblnLike = False Then mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", ""): mblnLike = True
        vPoint = GetCoordPos(txtInfo(txtʡ).hWnd, txtInfo(txtʡ).Left, txtInfo(txtʡ).Top)
        For i = 1 To intIndex
            If txtInfo(i - 1).Visible Then vPoint.X = vPoint.X + txtInfo(i - 1).Width
        Next
        Set rsTmp = zlDatabase.ShowSQLSelect(UserControl.Parent, strSQL, 0, "����", False, "", "", False, _
            False, True, vPoint.X, vPoint.Y, txtInfo(intIndex).Height, blnCancel, False, False, _
            UCase(strInput) & "%", mstrLike & UCase(strInput) & "%", intLevel, strCode)
        '������������,��һ��Ҫƥ��
        If Not rsTmp Is Nothing Then
            mblnEdit = True
            Call FillItems(intIndex, rsTmp, blnClare)
            strPreCode = rsTmp!�ϼ����� & ""
            '�Զ���ȱ
            strCode = strPreCode: strPreCode = ""
            For i = intIndex - 1 To 0 Step -1
                If Not marrItems(i).bln��Ч Then
                    Call SetInput(i, strCode, strPreCode, False)
                    strCode = strPreCode: strPreCode = ""
                    If strCode = "" Then Exit For
                End If
            Next
            If Not SetNoNaturalAd(intIndex) Then
                If intIndex = txt��ϸ��ַ Then
                    txtInfo(txt��ϸ��ַ).SelLength = 0
                    txtInfo(txt��ϸ��ַ).SelStart = 0
                    txtInfo(txt��ϸ��ַ).SelLength = Len(txtInfo(txt��ϸ��ַ).Text)
                    Exit Function
                End If
                Call LocateItem(intIndex, 1, True)
            End If
        Else
            mblnEdit = False
            Call zlControl.TxtSelAll(txtInfo(intIndex))
            If Not blnCancel And intIndex <> txt��ϸ��ַ Then
                If zlCommFun.IsCharChinese(txtInfo(intIndex).Text) Then
                    If MsgBox("�ֵ����δ�ҵ�������������Ƿ�Ҫʹ�������ֵ��", vbQuestion + vbYesNo + vbDefaultButton1, "��������") = vbYes Then
                        marrItems(intIndex).blnƥ�� = True
                        Call LocateItem(intIndex, 1, True)
                        mblnEdit = True
                    End If
                Else
                    MsgBox "�ֵ����δ�ҵ������������", vbInformation, "��������"
                    SetInput = True
                End If
            ElseIf intIndex <> txt��ϸ��ַ Or blnCancel Then
                Call LocateItem(intIndex, 0, True)
                Call FillItems(intIndex)
            Else
                marrItems(intIndex).blnƥ�� = True
                Call LocateItem(intIndex, 1, True)
                mblnEdit = True
            End If
        End If
        If Not rsTmp Is Nothing Then
            RaiseEvent SetInput(intLevel, rsTmp)
        End If
        RaiseEvent SetEdit(mblnEdit)
    Else
        'û����������Ҫƥ��
        strSQL = "select ����,����,�ϼ�����,�Ƿ�����,�Ƿ���ʾ from ���� where ����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����ȷ����", strInputCode & "")
        If Not rsTmp.EOF Then
            strPreCode = rsTmp!�ϼ����� & ""
        End If
        Call FillItems(intIndex, rsTmp, False)
    End If
End Function

Private Function SetNoNaturalAd(ByVal intIndex As Integer) As Boolean
'���ܣ��ж�һ����ַ���¼��Ƿ�û��ʵ�ʵ�ַ��ȫ�������ַ
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnOnlyVir As Boolean
    
    If Me.Items < Two Then Exit Function
    If intIndex = txt���� Then Exit Function
    If marrItems(intIndex).strCode <> "" And intIndex < Me.Items - 1 Then
        strSQL = "Select 1 ���� from ���� Where �ϼ����� =[1]  And Nvl(�Ƿ�����,0)=0 And Nvl(�Ƿ���ʾ,0)=0 And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ�û����Ч�¼�", marrItems(intIndex).strCode)
        If Me.Items = Two Then
            strSQL = "Select 1 ���� from ���� Where �ϼ����� =[1] And Rownum < 2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ�û����Ч�¼�", marrItems(intIndex).strCode)
            If rsTmp.RecordCount > 0 Then
                blnOnlyVir = True
            End If
        Else
            If rsTmp.RecordCount = 0 Then
                strSQL = "Select 1 ���� from ���� Where �ϼ����� =[1] And Rownum < 2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�ж��Ƿ�û����Ч�¼�", marrItems(intIndex).strCode)
                If rsTmp.RecordCount > 0 Then
                    blnOnlyVir = True
                End If
            End If
        End If
        If Not blnOnlyVir Then
            txtInfo(intIndex + 1).Locked = False
            If Me.Enabled And Not mblnLocked Then txtInfo(intIndex + 1).BackColor = vbWindowBackground
            marrItems(intIndex + 1).bln���� = False
            marrItems(intIndex + 1).bln����ʾ = False
             If Me.Items = Two Then  '����չʾʱ��������������ַ����ʾ��ַ,�������ƶ����е�λ��
                Call FillItems(txt��, , False)
             End If
             Call LocateItem(intIndex, 1, True)
        Else
            If Me.Enabled Then txtInfo(intIndex + 1).Locked = True
            If Me.ControlLock = False Then
                txtInfo(intIndex + 1).BackColor = vbButtonFace
            End If
            marrItems(intIndex + 1).bln���� = True
            marrItems(intIndex + 1).bln����ʾ = True
            If Me.Items = Two Then  '����չʾʱ��������������ַ����ʾ��ַ,�������ƶ����е�λ��
                Call FillItems(txt����, , False)
             End If
             Call LocateItem(intIndex + 1, 1, True)
        End If
        SetNoNaturalAd = True
    End If
End Function

Private Function StructAdress(ByVal strInput As String, Optional ByVal intType As Integer) As String
'���ܣ��ṹ����ַ������ȡ�ṹ����ַ��Ϣ��
    Dim rsTmp  As ADODB.Recordset, strSQL As String
    Dim arrAddress As Variant
    Dim arrTmp As Variant
    Dim i As Long, j As Long, blnClare As Boolean
    Dim blnCopyStruct As Boolean
    Dim str����Code As String
    
    If strInput Like "ZLSOFT:*" Then '�ṹ����ַ����
        If strInput Like "*,*,*,*,*,*|*,*,*,*,*,*|*,*,*,*,*,*|*,*,*,*,*,*|*,*,*,*,*,*" Then
            blnCopyStruct = True
        End If
        strInput = Mid(strInput, Len("ZLSOFT:") + 1)
    End If
    
    If strInput = "" Then
        strInput = ",,,,|,,,,|,,,,|,,,,|,,,,"
    ElseIf Not blnCopyStruct Then
        strSQL = "Select Zl_Adderss_Structure([1],[2]) ��ַ�ֽ� From dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ַ�ṹ�ֽ�", strInput, intType)
        strInput = rsTmp!��ַ�ֽ� & ""
    End If
    blnClare = True
    arrAddress = Split(strInput, "|")
    For i = LBound(arrAddress) To UBound(arrAddress)
        arrTmp = Split(arrAddress(i), ",")
        marrItems(i).strInfo = arrTmp(0)
        marrItems(i).strCode = arrTmp(1)
        marrItems(i).bln���� = Val(arrTmp(2)) = 1
        marrItems(i).bln����ʾ = Val(arrTmp(3)) = 1
        If UBound(arrTmp) = 5 Then '��ַ����ʱ���е����Ԫ��
            marrItems(i).blnƥ�� = Val(arrTmp(5)) = 1
        Else
            marrItems(i).blnƥ�� = True
        End If
        '�ü���ֻ�������¼�
        If Val(arrTmp(4)) = 1 Then
            marrItems(i).bln���� = True
            marrItems(i).bln����ʾ = True
        End If
        If Me.Items = Two And i = txt�� Then
            marrItems(i).bln���� = True
            marrItems(i).bln����ʾ = True
            marrItems(txt����).bln��Ч = Not marrItems(i).bln����
        End If
        If Me.Items = Two And i = txt���� And marrItems(txt����).strInfo = "" Then
            marrItems(i).strInfo = marrItems(txt��).strInfo
        End If
        If marrItems(i).bln��Ч Then
            If i = txt���� Then str����Code = marrItems(i).strCode
            If blnClare Then Call ClearItems(i, False): blnClare = False
        Else
             If i = txt��ϸ��ַ Then
                If Me.Items = Four Then '������ʾ��������ϲ���ϸ��ַ
                    marrItems(i).strInfo = marrItems(txt����).strInfo & marrItems(i).strInfo
                    '��ϸ��ַû�б��룬��ȡ����ı���
                    If marrItems(i).strCode = "" And str����Code <> "" Then
                        marrItems(i).strCode = str����Code
                        marrItems(i).strStName = marrItems(txt����).strInfo
                    End If
                    marrItems(txt����).strInfo = ""
                    Call FillItems(txt����, , False)
                End If
             End If
        End If
        Call FillItems(i, , False)
    Next
    If Me.value = marrItems(txt����).strInfo Then
        If marrItems(txt����).bln��Ч Then
            marrItems(txtʡ).strInfo = marrItems(txt����).strInfo
            marrItems(txt����).strInfo = ""
            Call FillItems(txtʡ, , False)
        End If
    End If
    If mblnChangeOld Then
        mstrOldAddress = Me.value
    End If
End Function

Private Function GetSetableItem(ByVal intIndex As Integer, Optional ByVal intStep As Integer) As Integer
'���ܣ����������������Ŀ
'������intIndex=��ʼ����
'         intStep=��λ����0-��ǰ��Ԫ��-1-��ǰ��λ����λ��ǰ�����һ��������ĵ�Ԫ��1-���λ����λ��������ĵ�Ԫ��
'���أ����Զ�λ�ĵ�Ԫ��
    Dim i As Integer, intReturn As Integer
    
    intReturn = -1
    If intStep = 0 Then
        '��ǰ��Ԫ���ܷ�λ�����ܶ�λ�������Ѱ��
        If Not (marrItems(intIndex).bln���� Or marrItems(intIndex).bln����ʾ) And txtInfo(intIndex).Visible Then
            intReturn = intIndex
        Else
            intReturn = GetSetableItem(intIndex, 1)
        End If
    ElseIf intStep = -1 Then '��ǰѰ��
        For i = intIndex - 1 To 0 Step -1
            If Not (marrItems(i).bln���� Or marrItems(i).bln����ʾ) And Not marrItems(i).bln��Ч Then intReturn = i: Exit For
        Next
    ElseIf intStep = 1 Then
        For i = intIndex + 1 To txt��ϸ��ַ
            If Not (marrItems(i).bln���� Or marrItems(i).bln����ʾ) And Not marrItems(i).bln��Ч Then intReturn = i: Exit For
        Next
    End If
    GetSetableItem = intReturn
End Function

Private Function LocateItem(ByVal intIndex As Integer, Optional ByVal intStep As Integer, Optional ByVal blnNotCheckSel As Boolean) As Integer
'���ܣ����ܶ�λ��Ŀ
'������intIndex=��ʼ����
'         intStep=��λ����0-��ǰ��Ԫ��-1-��ǰ��λ����λ��ǰ�����һ��������ĵ�Ԫ��1-���λ����λ��������ĵ�Ԫ��
'���أ����Զ�λ�ĵ�Ԫ��
    Dim intReturn As Integer
    Dim intStart As Integer, intEnd As Integer
    Dim i As Integer
    intReturn = -1
    If intStep = 0 Then
        intStart = intIndex
        intReturn = GetSetableItem(intIndex)
        intEnd = intReturn - 1
    ElseIf intStep = -1 Then
        intEnd = intIndex - 1
        If txtInfo(intIndex).SelStart = 0 Or blnNotCheckSel Then
            intReturn = GetSetableItem(intIndex, intStep)
        End If
        intStart = intReturn
    ElseIf intStep = 1 Then
        intStart = intIndex
        If txtInfo(intIndex).SelStart = Len(txtInfo(intIndex).Text) Or blnNotCheckSel Then
            intReturn = GetSetableItem(intIndex, intStep)
        End If
        intEnd = intReturn - 1
    End If
    If intReturn <> -1 Then
        txtInfo(intReturn).SetFocus
    ElseIf intStep >= 0 Then
        zlCommFun.PressKey (vbKeyTab)
        intEnd = txt��ϸ��ַ
    End If
    If intStart >= 0 Then
        For i = intStart To intEnd
            Call FillItems(i, , False)
        Next
    End If
End Function

Private Sub ClearItems(ByVal intIndex As Integer, Optional ByVal blnLocate As Boolean = True)
    Dim i As Integer, intEnd As Integer
    For i = intIndex To txt��ϸ��ַ
        marrItems(i).strCode = ""
        marrItems(i).bln���� = False
        marrItems(i).bln����ʾ = False
        If Me.Items = Two And i = txt�� Then
            marrItems(i).bln���� = True
            marrItems(i).bln����ʾ = True
        End If
        marrItems(i).blnƥ�� = False
        Call FillItems(i, , blnLocate And i = intIndex)
    Next
End Sub

Private Sub FillItems(ByVal intIndex As Integer, Optional ByVal rsInput As ADODB.Recordset, Optional ByVal blnLocate As Boolean = True)
    mblnSetItems = True
    '���ݼ�¼���������
    If Not rsInput Is Nothing Then
        marrItems(intIndex).strInfo = rsInput!���� & ""
        marrItems(intIndex).strCode = rsInput!���� & ""
        marrItems(intIndex).bln���� = Val(rsInput!�Ƿ����� & "") = 1
        If Me.Items = Two And intIndex = txt�� Then
            marrItems(intIndex).bln���� = True
        End If
        marrItems(intIndex).bln����ʾ = Val(rsInput!�Ƿ���ʾ & "") = 1
        marrItems(intIndex).blnƥ�� = True
    End If
    If intIndex > txt��ϸ��ַ Then Exit Sub
    '������ַ�ڶ���Ϊ����ʱ�����⴦��
    If Me.Items = Two Then
        marrItems(txt����).bln��Ч = Not marrItems(txt��).bln����
        If marrItems(txt����).bln��Ч Then
            If txtInfo(txt����).Visible Then marrItems(txt����).strInfo = ""
        Else
            If txtInfo(txt��).Visible Then marrItems(txt��).strInfo = ""
        End If
        Call UserControl_Resize
    End If
    '���������ݣ�������չʾ��ʽ
    txtInfo(intIndex).Text = marrItems(intIndex).strInfo
    txtInfo(intIndex).Tag = marrItems(intIndex).strCode
    If blnLocate Then
        If txtInfo(intIndex).Text = marrItems(intIndex).strNullInfo Then txtInfo(intIndex).Text = ""
        txtInfo(intIndex).ForeColor = &H80000008
    Else
        If txtInfo(intIndex).Text = "" Then txtInfo(intIndex).Text = marrItems(intIndex).strNullInfo
        If txtInfo(intIndex).Text = marrItems(intIndex).strNullInfo Then
            txtInfo(intIndex).ForeColor = mcolForeColor
            txtInfo(intIndex).SelStart = 0
            txtInfo(intIndex).SelLength = 0
        Else
            txtInfo(intIndex).ForeColor = &H80000008
        End If
    End If
    If marrItems(intIndex).bln��Ч Then
        txtInfo(intIndex).BackColor = vbButtonFace
    Else
        If marrItems(intIndex).bln���� Then
            If marrItems(intIndex).bln����ʾ Then
                txtInfo(intIndex).Text = ""
            End If
            txtInfo(intIndex).Enabled = Me.Enabled And Not marrItems(intIndex).bln����ʾ
            txtInfo(intIndex).Locked = Not Me.Enabled
            txtInfo(intIndex).BackColor = vbButtonFace
        Else
            txtInfo(intIndex).Enabled = Me.Enabled
            txtInfo(intIndex).Locked = mblnLocked Or Not Me.Enabled
            If Me.Enabled Then
                txtInfo(intIndex).BackColor = IIf(mblnLocked, vbButtonFace, Me.TextBackColor)
            Else
                txtInfo(intIndex).BackColor = vbButtonFace
            End If
        End If
    End If
    mblnSetItems = False
End Sub

Private Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Private Sub SetLine(ByVal lngStyle As Long)
'���ܣ������»���
    Dim i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    UserControl.Cls
    For i = 0 To txtInfo.Count - 1
        If lngStyle = Underline Then
            x1 = txtInfo(i).Left
            y1 = txtInfo(i).Top + txtInfo(i).Height
            x2 = txtInfo(i).Left + txtInfo(i).Width - 30
            y2 = y1
            UserControl.Line (x1, y1)-(x2, y2)
        End If
    Next
End Sub

Private Sub TxtMouseDown(ByRef ObjText As Object, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'���ܣ�TextBox��Ĭ���Ҽ���Ϣ��ʾ���޸�
    If intButton = vbRightButton Then
        If glngTXTProc = 0 Then
            glngTXTProc = GetWindowLong(ObjText.hWnd, GWL_WNDPROC)
            Call SetWindowLong(ObjText.hWnd, GWL_WNDPROC, AddressOf WndMessageMenu)
        End If
    End If
End Sub

Private Sub TxtMouseUp(ByRef ObjText As Object, ByRef intButton As Integer, ByRef intShift As Integer, ByRef sngX As Single, ByRef sngY As Single)
'���ܣ�TextBox��Ĭ���Ҽ���Ϣ��ʾ���޸�
    If intButton = vbRightButton Then
        If glngTXTProc <> 0 Then
            Call SetWindowLong(ObjText.hWnd, GWL_WNDPROC, glngTXTProc)
            glngTXTProc = 0
        End If
    End If
End Sub

Friend Sub PopMenu()
    Dim strTxt As String
    mnuPopuMenuCopyAll.Enabled = Me.value <> ""
    mnuPopuMenuDelete.Enabled = Me.value <> "" And Me.Enabled And Not Me.ControlLock
    strTxt = Clipboard.GetText
    mnuPopuMenuPasteAll.Enabled = strTxt <> "" And Me.Enabled And Not Me.ControlLock
    mnuPopuMenuPaste.Enabled = strTxt <> "" And Me.Enabled And Not Me.ControlLock
    PopupMenu mnuPopuMenu
End Sub

