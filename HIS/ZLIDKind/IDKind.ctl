VERSION 5.00
Begin VB.UserControl IDKind 
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   915
   KeyPreview      =   -1  'True
   ScaleHeight     =   570
   ScaleWidth      =   915
   ToolboxBitmap   =   "IDKind.ctx":0000
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   150
      ScaleHeight     =   300
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   150
      Width           =   350
   End
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   650
   End
End
Attribute VB_Name = "IDKind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'CBOList ��ʽ:����|ȫ��|ˢ����־|�����ID|���ų���|ȱʡ��־|....;��
Private Const DEF_CBOList = "��|��������￨|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0"
Private Const DEF_KeyShift = 2      '2��ʾ����control��
Private Const DEF_CboWidth = 1600
Private Const DEF_SmallStyle = False
Private mblnSmallStyle As Boolean
Private mstrIDKindStr As String
Private mintIDKind As Integer
Private mlngKeyShift As Long
Private mintCardLen As Integer   '��ǰ������
Private mintDefaultIDkindIdx As Integer 'ȱʡ�����
Private mintDefaultIDkindCardLen As Integer 'ȱʡ������
Private mcolIDKinds As Collection
Private WithEvents mobjParent As Form
Attribute mobjParent.VB_VarHelpID = -1
Public Event Click()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event ItemClick(Index As Integer)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2 'ǳ����
Private Const BDR_RAISEDINNER = &H4 'ǳ͹��
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '��͹��
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '���
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame������ʽ
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '��Frame������ʽ
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_SOFT = &H1000
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private mblnDo As Boolean
Private Sub RaisEffect(picBox As PictureBox, Optional intStyle As Integer, Optional strName As String = "")
'���ܣ���PictureBoxģ���3Dƽ�水ť
'������intStyle:0=ƽ��,-1=����,1=͹��
    
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        If strName <> "" Then
            .Cls
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
        
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        .Refresh
    End With
End Sub


Private Sub CboSetWidth(ByVal hWnd_combo As Long, ByVal lngWidth As Long)
'���ܣ�����Combo�ؼ������б�Ŀ��
'�˴��Ŀ�����������б�Ŀ�ȣ���������TWIPΪ��λ
    Const CB_SETDROPPEDWIDTH As Long = &H160

    SendMessage hWnd_combo, CB_SETDROPPEDWIDTH, lngWidth / Screen.TwipsPerPixelX, 0
End Sub

Private Sub mobjParent_KeyDown(KeyCode As Integer, Shift As Integer)
    If UserControl.Enabled Then
        If Shift = mlngKeyShift Then
            'keycode:96��С���̵�0,105��9
            If KeyCode > 95 And KeyCode < 106 Then
                IDKind = KeyCode - 96
            ElseIf KeyCode = 123 Then   'Ctrol+F12,������ִ�е�������Ķ�������Ӧ���ȼ�
                Call pic_Click
            End If
        End If
    End If
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "���ÿؼ��Ŀ�����"
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "StandardColor;��Ϊ"
Attribute Enabled.VB_UserMemId = -514
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
   UserControl.Enabled = NewValue
   cbo.Enabled = NewValue
   pic.Enabled = NewValue
   PropertyChanged "Enabled"
End Property

Public Property Get IDKindStr() As String
Attribute IDKindStr.VB_Description = "���ʶ�������б�,�������֮ǰ��;�ָ�,��������|ǰ�������Ϊ��ݱ�ʶ��ʾ,���ֱ�ʾ�Ƿ������²�����click�¼�"
    IDKindStr = mstrIDKindStr
End Property

Public Property Let IDKindStr(ByVal New_IDKindStr As String)
    If CboAddData(New_IDKindStr) Then
        mstrIDKindStr = New_IDKindStr
        PropertyChanged "IDKindStr"
    Else
        MsgBox "����ֵ�����Ϲ涨�ĸ�ʽ,�밴���¸�ʽ����:" & vbCrLf & "��|���￨��,0;ҽ|ҽ����,0"
        '�ָ�����
        mstrIDKindStr = DEF_CBOList
        PropertyChanged "IDKindStr"
        CboAddData (mstrIDKindStr)
    End If
End Property

Public Property Get KeyShift() As Long
Attribute KeyShift.VB_Description = "���ʿ�ݼ�����ϼ���һ����,��ο�KeyDown�¼���shift����"
    KeyShift = mlngKeyShift
End Property

Public Property Let KeyShift(ByVal New_KeyShift As Long)
    If New_KeyShift < 256 Then
        mlngKeyShift = New_KeyShift
        PropertyChanged "KeyShift"
    Else
        MsgBox "��Ч������ֵ,��ο�KeyDown�¼���shift����", vbInformation, App.ProductName
    End If
End Property

Public Property Get SmallStyle() As Boolean
    SmallStyle = mblnSmallStyle
End Property

Public Property Let SmallStyle(ByVal New_SmallSytle As Boolean)
    Dim MyFont As New StdFont
    
    mblnSmallStyle = New_SmallSytle
    If New_SmallSytle Then
        MyFont.Size = 10
    Else
        MyFont.Size = 12
    End If
    
    Set cbo.Font = MyFont
    Set pic.Font = MyFont
    pic.Height = cbo.Height - 60
    
    Call cbo_Click
End Property

'�����Խ�������ʱ��Ч
Public Property Get IDKind() As Integer
Attribute IDKind.VB_Description = "���õ�ǰ���������,��Ŵ�0��ʼ"
Attribute IDKind.VB_MemberFlags = "400"
   IDKind = mintIDKind
End Property

Public Property Let IDKind(ByVal New_IDKind As Integer)
   If New_IDKind <= UBound(Split(mstrIDKindStr, ";")) Then
        mintIDKind = New_IDKind
        PropertyChanged "IDKind"
        mblnDo = True
        cbo.ListIndex = New_IDKind
        mblnDo = False
   End If
End Property
Private Function CboAddData(ByVal strIDKindStr As String) As Boolean
    Dim arrTmp As Variant, i As Long
    Dim varKinds As Variant, j As Long
    On Error GoTo errHand
    If strIDKindStr = "" Then GoTo errHand
    cbo.Clear
    Set mcolIDKinds = New Collection
    mintDefaultIDkindIdx = -1: mintDefaultIDkindCardLen = -1
    varKinds = Split(strIDKindStr, ";")
    For i = 0 To UBound(varKinds)
        'CBOList ��ʽ:����|ȫ��|ˢ����־|�����ID|���ų���|ȱʡ��־|...;��
        If InStr(1, varKinds(i), "|") <= 0 Then GoTo errHand    'Or Len(varKinds(i)) > 20:���˺�ȡ��,ԭ����û��Ҫ�ж�
        arrTmp = Split(varKinds(i), "|")
        If UBound(arrTmp) <= 5 Then
            j = 5 - UBound(arrTmp)
            arrTmp = Split(varKinds(i) & String(j, "|"), "|")
        End If
        If Val(arrTmp(5)) = 1 Then
            'ȱʡ��־
            mintDefaultIDkindIdx = i
            mintDefaultIDkindCardLen = Val(arrTmp(4))
        End If
        mcolIDKinds.Add arrTmp, "K" & i
        cbo.AddItem Split(varKinds(i), "|")(1)
        cbo.ItemData(cbo.NewIndex) = Split(varKinds(i), "|")(2)
    Next
    cbo.ListIndex = 0   '����cbo_Click�¼�
    CboAddData = True
    Exit Function
errHand:
    CboAddData = False
End Function

Private Sub cbo_Click()
    If cbo.Locked Or cbo.ListIndex < 0 Then Exit Sub
    mintCardLen = Val(GetKindItem("���ų���"))
    
    pic.Height = cbo.Height - 60
    Call RaisEffect(pic, 1, CStr(mcolIDKinds("K" & cbo.ListIndex)(0)))
    If IDKind <> cbo.ListIndex Then
        IDKind = cbo.ListIndex  '�ֶ���������ťʱ,��Ҫ����idkind����
    End If
    RaiseEvent ItemClick(IDKind)
    'If cbo.ItemData(cbo.ListIndex) = 0 And cbo.Visible And Not mblnDo Then Call SendKeys("{Tab}"): Call SendKeys("{Tab}")
End Sub
Private Sub pic_Click()
    If cbo.ListIndex < 0 Then Exit Sub
    If cbo.ItemData(cbo.ListIndex) = 1 Then RaiseEvent Click
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cbo.ItemData(cbo.ListIndex) = 1 Then Call RaisEffect(pic, -1)
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cbo.ItemData(cbo.ListIndex) = 1 Then Call RaisEffect(pic, 1)
End Sub

Private Sub UserControl_EnterFocus()
    Call cbo.SetFocus
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub



'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
Private Sub UserControl_Initialize()
    pic.Left = 30
    cbo.Left = 0
    pic.Top = 30
    cbo.Top = 0
    UserControl.Width = cbo.Width
    UserControl.Height = cbo.Height
    Call CboSetWidth(cbo.hwnd, DEF_CboWidth)
End Sub

Private Sub UserControl_InitProperties()
    IDKindStr = DEF_CBOList
    KeyShift = DEF_KeyShift
    SmallStyle = DEF_SmallStyle
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo errH
    mblnSmallStyle = PropBag.ReadProperty("SmallStyle", DEF_SmallStyle)
    SmallStyle = mblnSmallStyle
    
    mlngKeyShift = PropBag.ReadProperty("KeyShift", DEF_KeyShift)
    mstrIDKindStr = PropBag.ReadProperty("IDKindStr", DEF_CBOList)
    Set cbo.Font = PropBag.ReadProperty("Font", Ambient.Font)
    pic.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Call CboAddData(mstrIDKindStr)
    If Ambient.UserMode Then Set mobjParent = UserControl.Parent
    Exit Sub
errH:
    MsgBox Err.Description, vbInformation, App.ProductName
    mblnSmallStyle = DEF_SmallStyle
    mlngKeyShift = DEF_KeyShift
    mstrIDKindStr = DEF_CBOList
    Call CboAddData(mstrIDKindStr)
    Set cbo.Font = PropBag.ReadProperty("Font", Ambient.Font)
    pic.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SmallStyle", mblnSmallStyle, DEF_SmallStyle)
    Call PropBag.WriteProperty("KeyShift", mlngKeyShift, DEF_KeyShift)
    Call PropBag.WriteProperty("IDKindStr", mstrIDKindStr, DEF_CBOList)
    Call PropBag.WriteProperty("Font", cbo.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", pic.ForeColor, &H80000012)
End Sub

Public Function IsMobileNo(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '����:�жϴ�����Ƿ�Ϊ�ֻ���
    '����:strinput-11λ�ֻ���
    '����:True-�������Ϊ�ֻ���;False-������벻Ϊ�ֻ���
    '����:������
    '����:2017-1-25
    '---------------------------------------------------------------------------------------------
    Dim strMobileRange As String
    If Not IsNumeric(strInput) Then Exit Function
    If Len(strInput) <> 11 Then Exit Function
    '�й��ƶ�
    strMobileRange = ",139,138,137,136,135,134,159,158,157,150,151,152,147,188,187,182,183,184,178"
    '�й���ͨ
    strMobileRange = strMobileRange & ",130,131,132,156,155,186,185,145,176"
    '�й�����
    strMobileRange = strMobileRange & ",133,153,189,180,181,177,173"
    '������Ӫ��
    strMobileRange = strMobileRange & ",170,"
    If InStr(strMobileRange, "," & Mid(strInput, 1, 3) & ",") = 0 Then Exit Function
    IsMobileNo = True
End Function

Private Sub UserControl_Resize()
    pic.Height = cbo.Height - 60
    UserControl.Width = cbo.Width
    UserControl.Height = cbo.Height
End Sub
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=cbo,cbo,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "����һ�� Font ����"
Attribute Font.VB_UserMemId = -512
    Set Font = cbo.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set cbo.Font = New_Font
    Set pic.Font = New_Font
    PropertyChanged "Font"
    Call UserControl_Resize
    Call cbo_Click
End Property
Public Property Get GetCardLength() As Integer
      GetCardLength = mintCardLen
End Property
Public Property Get ListCount() As Integer
      ListCount = cbo.ListCount
End Property

Public Property Get GetDefaultIDKindLength() As Integer
      '��ȡȱʡ�Ŀ��ų���
      GetDefaultIDKindLength = mintDefaultIDkindCardLen
End Property
Public Property Get GetDefaultIDKindIndex() As Integer
      '��ȡȱʡ����������
      GetDefaultIDKindIndex = mintDefaultIDkindIdx
End Property

Public Function GetKindIndex(ByVal strKindName As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ��KindName���Ƶ�����
    '���:strKindName-KindName����,Ҳ����Ϊ�����ID:���￨;
    '����:
    '����:����������ֵ
    '����:���˺�
    '����:2011-06-20 13:59:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 1 To mcolIDKinds.Count
        If IsNumeric(strKindName) Then
            '�����ID
            If Val(strKindName) = mcolIDKinds(i)(3) Then
                 GetKindIndex = i - 1: Exit Function
            End If
        ElseIf strKindName = mcolIDKinds(i)(1) Then
            GetKindIndex = i - 1: Exit Function
        End If
    Next
    GetKindIndex = -1
End Function

Public Function GetKindItem(ByVal strItemName As String, _
    Optional intKindIndex As Integer = -1) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡָ������Ŀֵ
    '���:strItemName-����Ϊָ������,Ҳ����Ϊ��������:
    '                               (����|ȫ��|ˢ����־|�����ID|���ų���|ȱʡ��־|....;��)
    '       KindIndex-���Ϊ��1,��ȡ��ǰ����
    '����:�����KindItemֵ
    '����:���˺�
    '����:2011-06-20 11:25:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    i = intKindIndex
    If intKindIndex <= -1 Then i = mintIDKind
    Err = 0: On Error Resume Next
    If IsNumeric(strItemName) Then
        j = Val(strItemName)
    '����|ȫ��|ˢ����־|�����ID|���ų���|������Ϣ;��
    ElseIf strItemName = "����" Then: j = 0
    ElseIf strItemName = "ȫ��" Then: j = 1
    ElseIf strItemName = "ˢ����־" Then: j = 2
    ElseIf strItemName = "�����ID" Then: j = 3
    ElseIf strItemName = "���ų���" Then: j = 4
    ElseIf strItemName = "ȱʡ��־" Then: j = 5
    End If
    Err = 0: On Error Resume Next
    If j >= 0 And j <= UBound(mcolIDKinds("K" & i)) Then
        GetKindItem = mcolIDKinds("K" & i)(j)
    Else
        GetKindItem = ""
    End If
    Err = 0: On Error GoTo 0
End Function
'ע�⣡��Ҫɾ�����޸����б�ע�͵��У�
'MappingInfo=pic,pic,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "����/���ö������ı���ͼ�ε�ǰ��ɫ��"
    ForeColor = pic.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    pic.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

