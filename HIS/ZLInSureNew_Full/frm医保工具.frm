VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmҽ������ 
   Caption         =   "ҽ������"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   Icon            =   "frmҽ������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   8940
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5490
      Left            =   2490
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5490
      ScaleWidth      =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   45
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   1950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":3A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":3C1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   5715
      Left            =   2550
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   10081
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   690
      Top             =   4380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":3E38
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":4152
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":446C
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":4786
            Key             =   "CommonD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":4AA0
            Key             =   "Disease"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   90
      Top             =   4380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":537A
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":5694
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":59AE
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":5CC8
            Key             =   "CommonD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":5FE2
            Key             =   "Disease"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":657C
            Key             =   "Limit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwKind_S 
      Height          =   5715
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   10081
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5715
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   635
      SimpleText      =   $"frmҽ������.frx":69CE
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmҽ������.frx":6A15
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10689
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmҽ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
Private mstrPrivs As String
Private Const clng�������� As Long = 1
Private Const clng����ҵ�� As Long = 2
Private Const clngסԺҵ�� As Long = 3
Private Const clng���� As Long = 4
Private Const clng���� As Long = 9998
Private Const clng�˳� As Long = 9000

Private mclsInsure As New clsInsure
Private rsMenu As New ADODB.Recordset       '�У��ϼ�ID,ID,����,Ȩ��
'��ģ��Ȩ���嵥:�ϴ������أ������������ã��������ݲ�ѯ�������սᣬ������㣬�������ݲ�ѯ����Ժ����Ժ�����㣬סԺ���ݲ�ѯ������

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strIcon As String, lst As ListItem
    On Error Resume Next
    
    Call RestoreWinState(Me, App.ProductName)
    
    mstrPrivs = gstrPrivs
    
    'װ�뱣�մ���
    gstrSQL = "select ���,����,�Ƿ�̶� from ������� where nvl(�Ƿ��ֹ,0)<>1 And ҽ������ Is NULL order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        '������ڴ����ʼ��ʱ���ã��Ͳ��ô�������������
        MsgBox "û�п��ñ�����𣬲���ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    lvwKind_S.ListItems.Clear
    Do Until rsTemp.EOF
        strIcon = IIf(rsTemp("�Ƿ�̶�") = 1, "Fix", "Common")
        Set lst = lvwKind_S.ListItems.Add(, "K" & rsTemp("���"), rsTemp("����"), strIcon, strIcon)
        
        rsTemp.MoveNext
    Loop
    If lvwKind_S.SelectedItem Is Nothing Then lvwKind_S.ListItems(1).Selected = True
    Call lvwKind_S_ItemClick(lvwKind_S.ListItems(1))
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    sngTop = 0
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    lvwKind_S.Top = sngTop
    lvwKind_S.Height = IIf(sngBottom - lvwKind_S.Top > 0, sngBottom - lvwKind_S.Top, 0)
    lvwKind_S.Left = ScaleLeft
    
    picSplitV.Top = sngTop
    picSplitV.Height = IIf(sngBottom - picSplitV.Top > 0, sngBottom - picSplitV.Top, 0)
    picSplitV.Left = lvwKind_S.Left + lvwKind_S.Width
    
    lvwMain.Top = sngTop
    lvwMain.Left = picSplitV.Left + 35
    lvwMain.Width = ScaleWidth - lvwMain.Left
    lvwMain.Height = picSplitV.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwKind_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call LoadTools
End Sub

Private Sub lvwMain_DblClick()
    Dim lngKey As Long, lngKey_�ϼ� As Long
    Dim strȨ�� As String           'Ϊ�ձ�ʾ������Ȩ�޿���
    Dim blnOwner As Boolean
    Dim lvwItem As ListItem
    
    If lvwMain.ListItems.Count = 0 Then Exit Sub
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    lngKey = CLng(Mid(lvwMain.SelectedItem.Key, 3))
    lngKey_�ϼ� = Val(lvwMain.SelectedItem.Tag)
    
    If lngKey = clng�˳� Then
        Unload Me
        Exit Sub
    End If
    
    If lngKey = clng���� Then
        '��ʾ��һ�������ݲ��˳�
        If lngKey_�ϼ� = 0 Then
            lngKey = 0
        Else
            rsMenu.Filter = "ID=" & lngKey_�ϼ�
            lngKey = rsMenu!�ϼ�ID
            rsMenu.Filter = 0
        End If
    End If
    
    'װ������
    rsMenu.Filter = "�ϼ�ID=" & lngKey
    If rsMenu.RecordCount = 0 Then
        rsMenu.Filter = 0
        Call ExecuteFuncs(lngKey)
        Exit Sub
    End If
    
    lvwMain.ListItems.Clear
    With rsMenu
        Do While Not .EOF
            '�ж��Ƿ�ӵ��Ȩ��
            blnOwner = True
            strȨ�� = Nvl(rsMenu!Ȩ��)
            If strȨ�� <> "" Then
                blnOwner = (InStr(1, ";" & mstrPrivs & ";", ";" & strȨ�� & ";") <> 0)
            End If
            
            If blnOwner Then
                Set lvwItem = lvwMain.ListItems.Add(, "K_" & rsMenu!ID, rsMenu!����, 1)
                lvwItem.Tag = Nvl(rsMenu!�ϼ�ID, 0)
            End If
            
            .MoveNext
        Loop
        .MoveFirst
        '���볣����˳�
        Set lvwItem = lvwMain.ListItems.Add(, "K_" & clng����, "����", 2)
        lvwItem.Tag = Nvl(rsMenu!�ϼ�ID, 0)
        Set lvwItem = lvwMain.ListItems.Add(, "K_" & clng�˳�, "�˳�", 3)
        lvwItem.Tag = Nvl(rsMenu!�ϼ�ID, 0)
        If lngKey = 0 Then
            lvwMain.ListItems("K_" & clng����).Ghosted = True
        Else
            lvwMain.ListItems("K_" & clng����).Ghosted = False
        End If
        .Filter = 0
    End With
End Sub

Public Sub ShowForm(ByVal frmParent As Object, ByVal intinsure As Integer)
    On Error Resume Next
    mintInsure = intinsure
    Me.Show , frmParent
End Sub

Public Sub InitInsure(ByVal intinsure As Integer)
    mintInsure = intinsure
End Sub

Public Sub LoadTools()
    Dim lngCounts As Long '��¼������
    If lvwKind_S.ListItems.Count = 0 Then Exit Sub
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    mintInsure = Mid(lvwKind_S.SelectedItem.Key, 2)
    lvwMain.ListItems.Clear

    '��ʼ��������¼
    Call Record_Init(rsMenu, "�ϼ�ID," & adDouble & ",18|ID," & adDouble & ",18|����," & adLongVarChar & ",100|Ȩ��," & adLongVarChar & ",100")
    
    lvwMain.ListItems.Add , "K_" & clng��������, "��������", 1
    Call Record_Add(rsMenu, "�ϼ�ID|ID|����", "0|" & clng�������� & "|" & "��������")
    lvwMain.ListItems.Add , "K_" & clng����ҵ��, "����ҵ��", 1
    Call Record_Add(rsMenu, "�ϼ�ID|ID|����", "0|" & clng����ҵ�� & "|" & "����ҵ��")
    lvwMain.ListItems.Add , "K_" & clngסԺҵ��, "סԺҵ��", 1
    Call Record_Add(rsMenu, "�ϼ�ID|ID|����", "0|" & clngסԺҵ�� & "|" & "סԺҵ��")
    lvwMain.ListItems.Add , "K_" & clng����, "����", 1
    Call Record_Add(rsMenu, "�ϼ�ID|ID|����", "0|" & clng���� & "|" & "����")
    
    lvwMain.ListItems.Add , "K_" & clng����, "����", 2
    lvwMain.ListItems.Add , "K_" & clng�˳�, "�˳�", 3
    
    '��һ�㲻����ִ���˵���һ��Ĺ���
    lvwMain.ListItems("K_" & clng����).Ghosted = True
    Select Case mintInsure
    Case TYPE_��ͨ
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng�������� & "|100|" & "���ط�����Ŀ|����")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng�������� & "|101|" & "���ز�����Ŀ|����")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng�������� & "|102|" & "���ر�׼ҩƷĿ¼|����")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng�������� & "|103|" & "����ҩƷִ�п�|����")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng����ҵ�� & "|200|" & "�������|�����ս�")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng����ҵ�� & "|201|" & "��ҩ��ϸ��ѯ|")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng����ҵ�� & "|202|" & "�ֹ������������|�������")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clngסԺҵ�� & "|300|" & "סԺ�����ѯ(��סԺ��¼,������ϸ,���û���)|")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clngסԺҵ�� & "|301|" & "ҽ����Ժ����|��Ժ")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng���� & "|400|" & "����Ա����|������������")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng���� & "|401|" & "�걨ҩƷ|������������")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng���� & "|402|" & "����Ա��λ|������������")
    Case TYPE_����
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����", clng����ҵ�� & "|200|" & "�ֹ������������")
    Case TYPE_ͭɽ��
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clngסԺҵ�� & "|300|" & "����ϴ���־|")
    Case TYPE_��������
        Call ҽ����ʼ��_��������
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng�������� & "|100|" & "���������ϴ�|�ϴ�")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng�������� & "|101|" & "����ҽ������Ŀ¼|����")
'        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng�������� & "|102|" & "��ĿĿ¼ά��|")
    Case TYPE_������
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng�������� & "|100|" & "���㵥����|�ϴ�")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng����ҵ�� & "|200|" & "����תҽ��|�������")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng����ҵ�� & "|201|" & "����������ݺ˶�|�������")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng����ҵ�� & "|202|" & "ҽ������������|�������")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clngסԺҵ�� & "|300|" & "����ҩƷ����|סԺ���ݲ�ѯ")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clngסԺҵ�� & "|301|" & "�����ϴ�����|�ϴ�")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng���� & "|400|" & "��������δ�ɹ������Ľ������|������������")
        Call Record_Add(rsMenu, "�ϼ�ID|ID|����|Ȩ��", clng���� & "|401|" & "����������ҩ����|������������")
    End Select
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With picSplitV
        .Move .Left + x
    End With
    Me.lvwKind_S.Width = picSplitV.Left
    Call Form_Resize
End Sub

Private Function ExecuteFuncs(ByVal lng���ܺ� As Long) As Boolean
    Dim blnDo As Boolean
    On Error GoTo errHand
    
    If lng���ܺ� < 100 Then Exit Function
    Select Case mintInsure
    Case TYPE_��ͨ
        blnDo = ��ͨ���߰�(lng���ܺ�)
    Case TYPE_����
        blnDo = ���깤�߰�(lng���ܺ�)
    Case TYPE_ͭɽ��
        blnDo = ͭɽ��ҽ�����߰�(lng���ܺ�)
    Case TYPE_��������
        blnDo = ���󹤾߰�(lng���ܺ�)
    Case TYPE_������
        blnDo = �������߰�(lng���ܺ�)
    End Select
    
    ExecuteFuncs = blnDo
    If blnDo Then MsgBox "ִ�гɹ���", vbInformation, gstrSysName
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function �������߰�(ByVal lng���ܺ� As Long) As Boolean
    Dim lngID As Long
    Dim str���� As String, str���� As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    On Error GoTo errHand
    
    If Not gclsInsure.InitInsure(gcnOracle, TYPE_������) Then Exit Function
    
    Select Case lng���ܺ�
        Case 100
            Call frm���㵥����_����.ShowME(mintInsure)
        Case 200
            Call frm����תҽ��.ShowME(mintInsure)
        Case 201
            Call frm����������ݺ˶�_����.ShowME(1, mintInsure)
        Case 202
            With frmIdentify����������
                .intinsure = mintInsure
                .Show vbModal
            End With
            Set frmIdentify���������� = Nothing
        Case 101
            '���浽���ǵĲ���Ŀ¼����
            If InitXML = False Then Exit Function
            If Not CommServer("QUERYSPECILLNESS") Then Exit Function
            Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
            If nodRowset Is Nothing Then Exit Function
            
            '���ݱ���õ���������
            For Each nodRow In nodRowset.childNodes
                lngID = zlDatabase.GetNextID("���ղ���")
                str���� = GetAttributeValue(nodRow, "SPECILLNESSCODE")
                str���� = GetAttributeValue(nodRow, "SPECILLNESSNAME")
                gstrSQL = "zl_���ղ���_INSERT(" & lngID & "," & TYPE_������ & ",'" & str���� & "','" & str���� & "','" & zlCommFun.SpellCode(str����) & "',2,0,0)"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "���²���")
            Next
        Case 300
            Call frm������Ŀ����_����.ShowSelect(mintInsure)
        Case 301
            With frm���ս����ϴ�_����
                .Insure = TYPE_������
                .Show vbModal
            End With
            Set frm���ս����ϴ�_���� = Nothing
        Case 400
            With frmIdentify�����������
                .Insure = mintInsure
                .Show vbModal
            End With
            Set frmIdentify����������� = Nothing
        Case 401
            frm����ҩƷ����_����.ShowME mintInsure
    End Select
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ���󹤾߰�(ByVal lng���ܺ� As Long) As Boolean
    If Not gclsInsure.InitInsure(gcnOracle, TYPE_��������) Then Exit Function
    
    Select Case lng���ܺ�
        Case 100    '���������ϴ�
            frmMain_�������󲡰���Ϣ.Show 1, Me
        Case 101
            With frmMain_��������Ŀ¼�·�
                .intinsure = TYPE_��������
                .Show 1, Me
            End With
        Case 102
            With frmMain_���������Ŀ����
                .Show 1, Me
            End With
    End Select
End Function

Private Function ͭɽ��ҽ�����߰�(ByVal lng���ܺ� As Long) As Boolean
    Const cͭɽ��_����ϴ���־ As Integer = 300
    Dim rsTmp As New ADODB.Recordset
    Dim lng����ID As Long, lng��ҳID As Long
    Select Case lng���ܺ�
    Case cͭɽ��_����ϴ���־
    
        gstrSQL = "Select a.����id||'_'||a.��ҳid as ID,a.����id, a.��ҳid, b.סԺ��, b.סԺ����, b.����, b.�Ա�, b.����, b.���֤��, a.��Ժ����" & vbNewLine & _
                "From ������ҳ a, ������Ϣ b" & vbNewLine & _
                "Where a.����id = b.����id And a.��ҳid = Nvl(b.סԺ����, 0) And b.��Ժ=1 And A.����=" & TYPE_ͭɽ��
        Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "ѡ����", True)
        If rsTmp Is Nothing Then
            MsgBox "����Ժҽ�����˿ɹ�ѡ��", vbQuestion, gstrSysName
        Else
            If rsTmp.State = 0 Then
                MsgBox "����Ժҽ�����˿ɹ�ѡ��", vbQuestion, gstrSysName
            Else
                If rsTmp.RecordCount > 0 Then
                    lng����ID = Nvl(rsTmp.Fields("����ID"), 0)
                    lng��ҳID = Nvl(rsTmp.Fields("��ҳID"), 0)
                    If MsgBox("��Ҫ���[" & Nvl(rsTmp.Fields("����")) & "]��סԺ��Ϊ��" & Nvl(rsTmp.Fields("סԺ��")) & "���ķ����ϴ���־����ȷ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        gstrSQL = "Update ���˷��ü�¼ Set �Ƿ��ϴ�=0 Where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID & " & And �Ƿ��ϴ�=1"
                        gcnOracle.Execute gstrSQL
                    End If
                End If
            End If
        End If
    End Select
End Function

Private Function ��ͨ���߰�(ByVal lng���ܺ� As Long) As Boolean
    Const C��ͨ_���ط�����Ŀ As Integer = 100
    Const C��ͨ_���ز�����Ŀ As Integer = 101
    Const C��ͨ_���ر�׼ҩƷĿ¼ As Integer = 102
    Const C��ͨ_����ҩƷִ�п� As Integer = 103
    Const C��ͨ_������� As Integer = 200
    Const C��ͨ_��ҩ��ϸ As Integer = 201
    Const C��ͨ_�ֹ��������ﵥ�� As Integer = 202
    Const C��ͨ_סԺ��� As Integer = 300
    Const C��ͨ_ҽ����Ժ���� As Integer = 301
    Const C��ͨ_����Ա���� As Integer = 400
    Const C��ͨ_�걨ҩƷ As Integer = 401
    Const C��ͨ_����Ա��λ As Integer = 402
    
    If lng���ܺ� <> C��ͨ_����Ա���� Then
        If Not mclsInsure.InitInsure(gcnOracle, TYPE_��ͨ) Then Exit Function
    Else
        If Not ҽ����ʼ��_��ͨ(False) Then Exit Function
    End If
    
    Select Case lng���ܺ�
    Case C��ͨ_�������
        ��ͨ���߰� = frmConn��ͨ.Execute("I250", 0, "", "���ڽ���ҽ���������......")
        Call ShowWindow(frmConn��ͨ.hwnd, 0)
    Case C��ͨ_��ҩ��ϸ
        ��ͨ���߰� = True
        Call frm��ͨ��ѯ����.ShowForm("��ҩ��ϸ��ѯ", "������ˮ��")
    Case C��ͨ_סԺ���
        ��ͨ���߰� = True
        Call frm��ͨ��ѯ����.ShowForm("סԺ�����ѯ", "סԺ��ˮ��")
    Case C��ͨ_���ط�����Ŀ
        ��ͨ���߰� = ��ͨ_���ط�����Ŀ
    Case C��ͨ_���ز�����Ŀ
        ��ͨ���߰� = ��ͨ_���ز�����Ŀ
    Case C��ͨ_���ر�׼ҩƷĿ¼
        ��ͨ���߰� = ��ͨ_���ر�׼ҩƷĿ¼
    Case C��ͨ_����ҩƷִ�п�
        ��ͨ���߰� = ��ͨ_����ҩƷִ�п�
    Case C��ͨ_����Ա����
        '����Ա�����ؽ���ҽ����ʼ��
        ��ͨ���߰� = ��ͨ_����Ա����
    Case C��ͨ_����Ա��λ
        ��ͨ���߰� = ��ͨ_����Ա��λ
    Case C��ͨ_�ֹ��������ﵥ��
        ��ͨ���߰� = ��ͨ_�ֹ��������ﵥ��
    Case C��ͨ_�걨ҩƷ
        ��ͨ���߰� = ��ͨ_�걨ҩƷ
    Case C��ͨ_ҽ����Ժ����
        Dim StrInput As String
        Dim rsTemp As New ADODB.Recordset
        On Error GoTo errHand
        StrInput = InputBox("������ò��˵�HISסԺ�ţ�", "ҽ����Ժ����")
        If Trim(StrInput) = "" Then Exit Function
        
        gstrSQL = " Select A.˳��� From �����ʻ� A,������Ϣ B Where A.����ID=B.����ID And B.סԺ��=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ��סԺ��", StrInput)
        If rsTemp.RecordCount = 0 Then Exit Function
        If frmConn��ͨ.Execute("I345", 0, Nvl(rsTemp!˳���), "���ڽ���ҽ����Ժ����......") = False Then Exit Function
        MsgBox "ҽ����Ժ�����ɹ����������°����Ժ����ҵ���ˣ�", vbInformation, gstrSysName
        Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
        Exit Function
    End Select
End Function

Private Function ��ͨ_����Ա����() As Boolean
    frmUserList.Show vbModal
End Function

Private Function ��ͨ_����Ա��λ() As Boolean
    Dim strCode As String
    strCode = InputBox("���������λ����Ա��ţ�", "����Ա��λ", "00")
    If Trim(strCode) = "" Then Exit Function
    If Len(strCode) > 2 Then Exit Function
    
    If Not frmConn��ͨ.Execute("I050", 3, strCode, "����Ա��λ......") Then Exit Function
    MsgBox "����Ա��λ�ɹ���", vbInformation, gstrSysName
End Function

Private Function ��ͨ_�ֹ��������ﵥ��() As Boolean
    Dim arrData
    Dim strData As String
    
    strData = InputBox("�����������������ʻ����м���;�ŷָ���", "�ֹ��������ﵥ��", "1010111;50.00")
    If Trim(strData) = "" Then Exit Function
    If InStr(1, strData, ";") = 0 Then
        MsgBox "��ʽ������ȷ��ʽ��������;�����ʻ����磺1010111;50.00", vbInformation, gstrSysName
        Exit Function
    End If
    If Not IsNumeric(Split(strData, ";")(1)) Then
        MsgBox "��ʽ������ȷ��ʽ��������;�����ʻ����磺1010111;50.00", vbInformation, gstrSysName
        Exit Function
    End If
    arrData = Split(strData, ";")
    
    If Not frmConn��ͨ.Execute("I220", 0, arrData(0) & vbTab & arrData(1), "�ֹ��������ﵥ��......") Then Exit Function
    MsgBox "�ֹ��������ﵥ�ݳɹ���", vbInformation, gstrSysName
End Function

Private Function ��ͨ_�걨ҩƷ() As Boolean
    '��HIS����ȡ��׼Ŀ¼�е����ݽ�����Ŀ���գ�Ȼ��ͳһ�걨���ɹ������ִ�п�
    Dim rsTemp As New ADODB.Recordset, rs��Ŀ As New ADODB.Recordset, strTemp As String
    
    Set rs��Ŀ = gcnOracle.Execute("Select a.�շ�ϸĿid,a.��Ŀ����,b.���㵥λ,b.����,b.��� From ����֧����Ŀ a,�շ�ϸĿ b Where ����=103 and nvl(��ע,'δ����')='δ����' And a.�շ�ϸĿid=b.id And (b.���='5' or b.���='6' or b.���='7')")
    If rs��Ŀ.EOF Then
        MsgBox "û����Ŀ��Ҫ�걨", vbInformation, "�걨"
        Exit Function
    End If
    
    While Not rs��Ŀ.EOF
        Set rsTemp = gcnOracle.Execute("Select * from ҩƷĿ¼ Where ҩƷID=" & rs��Ŀ!�շ�ϸĿID)
        strTemp = LeftStr(rs��Ŀ!����, 28) & vbTab & LeftStr(Nvl(rs��Ŀ!���㵥λ, " "), 4) & _
            vbTab & LeftStr(Nvl(rsTemp!���, " "), 16) & vbTab & _
            LeftStr(Nvl(rsTemp!ҩƷ��Դ, " "), 12) & vbTab & rsTemp!ָ�����ۼ� & vbTab & _
            LeftStr(Nvl(rsTemp!����, " "), 30) & vbTab & LeftStr(Nvl(rsTemp!��׼�ĺ�, " "), 30)
        Set rsTemp = gcn��ͨ.Execute("Select * From tab_syml Where dm='" & rs��Ŀ!��Ŀ���� & "'")
        strTemp = rs��Ŀ!�շ�ϸĿID & vbTab & " " & vbTab & rsTemp!dm & vbTab & _
            rsTemp!lb & vbTab & strTemp
        
        frmConn��ͨ.Execute "I110", 1, strTemp, "���ڽ���ҩƷ�걨......"
        rs��Ŀ.MoveNext
    Wend
    ��ͨ_�걨ҩƷ = True
End Function

Private Function ��ͨ_����ҩƷִ�п�() As Boolean
    '����ҩƷĿ¼ִ�п�
    Dim str����״̬ As String, rsTemp As New ADODB.Recordset, strData() As String, lngLoop As Long
    Dim strTemp As String
    On Error GoTo errHandle

    '��ǰ�ķ�ʽ����ȡ����Ϊʹ��11���ܺţ��Ƚ�������Ŀ����Ϊδ��ˣ��ٸ��ݷ������ݸ��£�ֻҪ�оͿ϶���������
    gcnOracle.BeginTrans
    If Not frmConn��ͨ.Execute("I110", 11, "", "���ڻ�ȡҩƷĿ¼ִ�п�����......") Then gcnOracle.RollbackTrans: Exit Function

    Call ShowWindow(frmConn��ͨ.hwnd, 9)
    DoEvents

    gstrSQL = "Update ����֧����Ŀ " & _
             " Set ��ע='δ����' " & _
             " Where ����=103 And �Ƿ�ҽ��=1 And ��Ŀ���� is not null " & _
             " And �շ�ϸĿID In ( " & _
             "     Select Id From �շ�ϸĿ Where ��� In ('5','6','7'))"
    gcnOracle.Execute gstrSQL
    For lngLoop = 1 To frmConn��ͨ.mlngRows
        If frmConn��ͨ.Query(lngLoop - 1, 1, "���ڸ�������(" & lngLoop & "/" & frmConn��ͨ.mlngRows & ")......") = False Then gcnOracle.RollbackTrans: Exit Function
        strTemp = Replace(frmConn��ͨ.strReturnInfo, "'", "''")
        gcnOracle.Execute "update ����֧����Ŀ set ��ע='������' Where �շ�ϸĿID=" & Split(strTemp, vbTab)(0)
    Next
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
    
    gcnOracle.CommitTrans
    ��ͨ_����ҩƷִ�п� = True
    Exit Function

errHandle:
    If MsgBox("������Ŀʱ��������" & vbCrLf & Err.Description & vbCrLf & "�Ƿ����ԣ�", vbInformation + vbRetryCancel, "����") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
    gcnOracle.RollbackTrans
End Function

Private Function ��ͨ_���ر�׼ҩƷĿ¼() As Boolean
    '���ر�׼ҩƷĿ¼
    Dim str����״̬ As String, rsTemp As New ADODB.Recordset, strData() As String, lngLoop As Long
    Dim strTemp As String
    Dim arrData
    
    If Not frmConn��ͨ.Execute("I100", 0, "", "��ȡҽ�����ݸ���״̬......") Then Exit Function
    If frmConn��ͨ.Query(0, 0) = False Then Exit Function
    If frmConn��ͨ.strReturnInfo = "" Then
        If MsgBox("����ȡ��ҽ�����ݸ���״̬���Ƿ������", vbQuestion + vbYesNo, "����") = vbNo Then
            Exit Function
        Else
            str����״̬ = ""
        End If
    Else
        str����״̬ = Split(frmConn��ͨ.strReturnInfo, vbTab)(0)
    End If
    
    Set rsTemp = gcn��ͨ.Execute("Select * From TAB_UPDATE")
    If rsTemp.EOF Then
        gcn��ͨ.Execute "Insert Into TAB_UPDATE Values (NULL,'" & str����״̬ & "',NULL)"
    Else
        If IsNull(rsTemp!BZML) Then
            gcn��ͨ.Execute "Update TAB_UPDATE Set BZML='" & str����״̬ & "'"
        ElseIf rsTemp!BZML = str����״̬ Then
            If MsgBox("���ϴ�������������׼ҩƷĿ¼δ���и��£��Ƿ��������أ�", vbYesNo + vbQuestion, "���ر�׼ҩƷĿ¼") = vbNo Then
                Exit Function
            End If
        Else
            gcn��ͨ.Execute "Update TAB_UPDATE Set BZML='" & str����״̬ & "'"
        End If
    End If
    
    If Not frmConn��ͨ.Execute("I100", 3, "", "���ڻ�ȡ��׼ҩƷĿ¼����......") Then Exit Function
    
    On Error GoTo errHandle
    gcn��ͨ.BeginTrans
    gcn��ͨ.Execute "Delete From tab_syml"
    Call ShowWindow(frmConn��ͨ.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn��ͨ.mlngRows
        If frmConn��ͨ.Query(lngLoop - 1, 1, "���ڸ�������(" & lngLoop & "/" & frmConn��ͨ.mlngRows & ")......") = False Then
            gcn��ͨ.RollbackTrans
            Exit Function
        End If
        strTemp = frmConn��ͨ.strReturnInfo
        arrData = Split(strTemp, vbTab)
        strTemp = Replace(arrData(11), "'", "") 'rq
        If Trim(strTemp) <> "" Then
            strTemp = "to_date('" & strTemp & "','yyyyMMdd')"
        Else
            strTemp = "''"
        End If
        gcn��ͨ.Execute "Insert Into tab_syml (dl,ty,dm,tm,sm,lb,py,dw,dj,jx,gg,rq,zt,xd,xj,cs) " & _
            " values (" & _
            "'" & Replace(arrData(0), "'", "''") & "','" & Replace(arrData(1), "'", "''") & "'," & _
            "'" & Replace(arrData(2), "'", "''") & "','" & Replace(arrData(3), "'", "''") & "'," & _
            "'" & Replace(arrData(4), "'", "''") & "','" & Replace(arrData(5), "'", "''") & "'," & _
            "'" & Replace(arrData(6), "'", "''") & "','" & Replace(arrData(7), "'", "''") & "'," & _
            "'" & Replace(arrData(8), "'", "''") & "','" & Replace(arrData(9), "'", "''") & "'," & _
            "'" & Replace(arrData(10), "'", "''") & "'," & strTemp & "," & _
            "'" & Replace(arrData(12), "'", "''") & "','" & Replace(arrData(13), "'", "''") & "'," & _
            "'" & Replace(arrData(14), "'", "''") & "','" & Replace(arrData(15), "'", "''") & "')"
    Next
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
    gcn��ͨ.CommitTrans
    ��ͨ_���ر�׼ҩƷĿ¼ = True
    Exit Function
    
errHandle:
    If MsgBox("������Ŀʱ��������" & vbCrLf & Err.Description & vbCrLf & "�Ƿ����ԣ�", vbInformation + vbRetryCancel, "����") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
    gcn��ͨ.RollbackTrans
End Function

Private Function ��ͨ_���ز�����Ŀ() As Boolean
    '���ز�����Ŀ
    Dim str����״̬ As String, rsTemp As New ADODB.Recordset, strData() As String, lngLoop As Long
    Dim strTemp As String
    
    If Not frmConn��ͨ.Execute("I100", 0, "", "��ȡҽ�����ݸ���״̬......") Then Exit Function
    If frmConn��ͨ.Query(0, 0) = False Then Exit Function
    
    If frmConn��ͨ.strReturnInfo = "" Then
        If MsgBox("����ȡ��ҽ�����ݸ���״̬���Ƿ������", vbQuestion + vbYesNo, "����") = vbNo Then
            Exit Function
        Else
            str����״̬ = ""
        End If
    Else
        str����״̬ = Split(frmConn��ͨ.strReturnInfo, vbTab)(0)
    End If
    
    Set rsTemp = gcn��ͨ.Execute("Select * From TAB_UPDATE")
    If rsTemp.EOF Then
        gcn��ͨ.Execute "Insert Into TAB_UPDATE Values (NULL,'" & str����״̬ & "',NULL)"
    Else
        If IsNull(rsTemp!CLXM) Then
            gcn��ͨ.Execute "Update TAB_UPDATE Set CLXM='" & str����״̬ & "'"
        ElseIf rsTemp!CLXM = str����״̬ Then
            If MsgBox("���ϴ�����������������Ŀδ���и��£��Ƿ��������أ�", vbYesNo + vbQuestion, "���ز�����Ŀ") = vbNo Then
                Exit Function
            End If
        Else
            gcn��ͨ.Execute "Update TAB_UPDATE Set CLXM='" & str����״̬ & "'"
        End If
    End If
    
    If Not frmConn��ͨ.Execute("I100", 2, "", "���ڻ�ȡ������Ŀ����......") Then Exit Function
'    If frmConn��ͨ.Query(0, 0) = False Then Exit Sub
'
'    strData = Split(frmConn��ͨ.strReturnInfo, Chr(10))
    
    On Error GoTo errHandle
    gcn��ͨ.BeginTrans
    gcn��ͨ.Execute "Delete From tab_fwcl Where lb In (31,32,33,51,52,53)"
    Call ShowWindow(frmConn��ͨ.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn��ͨ.mlngRows
        If frmConn��ͨ.Query(lngLoop - 1, 1, "���ڸ�������(" & lngLoop & "/" & frmConn��ͨ.mlngRows & ")......") = False Then
            gcn��ͨ.RollbackTrans
            Exit Function
        End If
        strTemp = frmConn��ͨ.strReturnInfo
        gcn��ͨ.Execute "Insert Into tab_fwcl (dm,kc,lb,mc,dw,dj,cx) values ('" & _
            Split(strTemp, vbTab)(0) & "','" & _
            Split(strTemp, vbTab)(1) & "'," & _
            Split(strTemp, vbTab)(2) & ",'" & _
            Split(strTemp, vbTab)(3) & "','" & _
            Split(strTemp, vbTab)(4) & "'," & _
            Split(strTemp, vbTab)(5) & "," & _
            IIf(Split(strTemp, vbTab)(6) = " ", "NULL", Split(strTemp, vbTab)(6)) & ")"
    Next
    Call ShowWindow(frmConn��ͨ.hwnd, 0)

    gcn��ͨ.CommitTrans
    ��ͨ_���ز�����Ŀ = True
    Exit Function
    
errHandle:
    If MsgBox("������Ŀʱ��������" & vbCrLf & Err.Description & vbCrLf & "�Ƿ����ԣ�", vbInformation + vbRetryCancel, "����") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
    gcn��ͨ.RollbackTrans
End Function

Private Function ��ͨ_���ط�����Ŀ() As Boolean
    '���ط�����Ŀ
    Dim str����״̬ As String, rsTemp As New ADODB.Recordset, strData() As String, lngLoop As Long
    Dim strTemp As String
    If Not frmConn��ͨ.Execute("I100", 0, "", "��ȡҽ�����ݸ���״̬......") Then Exit Function
    If frmConn��ͨ.Query(0, 0) = False Then Exit Function
    If frmConn��ͨ.strReturnInfo = "" Then
        If MsgBox("����ȡ��ҽ�����ݸ���״̬���Ƿ������", vbQuestion + vbYesNo, "����") = vbNo Then
            Exit Function
        Else
            str����״̬ = ""
        End If
    Else
        str����״̬ = Split(frmConn��ͨ.strReturnInfo, vbTab)(0)
    End If
    
    Set rsTemp = gcn��ͨ.Execute("Select * From TAB_UPDATE")
    If rsTemp.EOF Then
        gcn��ͨ.Execute "Insert Into TAB_UPDATE Values ('" & str����״̬ & "',NULL,NULL)"
    Else
        If IsNull(rsTemp!FWXM) Then
            gcn��ͨ.Execute "Update TAB_UPDATE Set FWXM='" & str����״̬ & "'"
        ElseIf rsTemp!FWXM = str����״̬ Then
            If MsgBox("���ϴ�����������������Ŀδ���и��£��Ƿ��������أ�", vbYesNo + vbQuestion, "���ط�����Ŀ") = vbNo Then
                Exit Function
            End If
        Else
            gcn��ͨ.Execute "Update TAB_UPDATE Set FWXM='" & str����״̬ & "'"
        End If
    End If
    
    If Not frmConn��ͨ.Execute("I100", 1, "", "���ڻ�ȡ������Ŀ����......") Then Exit Function
'    If frmConn��ͨ.Query(0, 0) = False Then Exit Sub
'    strData = Split(frmConn��ͨ.strReturnInfo, Chr(10))
    
    On Error GoTo errHandle
    gcn��ͨ.BeginTrans
    gcn��ͨ.Execute "Delete From tab_fwcl Where lb In (20,21,22,23,24,25,40)"
    Call ShowWindow(frmConn��ͨ.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn��ͨ.mlngRows
        If frmConn��ͨ.Query(lngLoop - 1, 1, "���ڸ�������(" & lngLoop & "/" & frmConn��ͨ.mlngRows & ")......") = False Then
            gcn��ͨ.RollbackTrans
            Exit Function
        End If
        strTemp = frmConn��ͨ.strReturnInfo
        gcn��ͨ.Execute "Insert Into tab_fwcl (dm,kc,lb,mc,dw,dj,cx) values ('" & _
            Split(strTemp, vbTab)(0) & "','" & _
            Split(strTemp, vbTab)(1) & "'," & _
            Split(strTemp, vbTab)(2) & ",'" & _
            Replace(Split(strTemp, vbTab)(3), "'", "") & "','" & _
            Split(strTemp, vbTab)(4) & "'," & _
            Split(strTemp, vbTab)(5) & "," & _
            IIf(Split(strTemp, vbTab)(6) = " ", "NULL", Split(strTemp, vbTab)(6)) & ")"
    Next
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
    gcn��ͨ.CommitTrans
    ��ͨ_���ط�����Ŀ = True
    Exit Function
    
errHandle:
    If MsgBox("������Ŀʱ��������" & vbCrLf & Err.Description & vbCrLf & "�Ƿ����ԣ�", vbInformation + vbRetryCancel, "����") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn��ͨ.hwnd, 0)
    gcn��ͨ.RollbackTrans
End Function

Private Function ���깤�߰�(ByVal lng���ܺ� As Long) As Boolean
    Dim strJZBH As String
    Const C����_�ֹ������������ As Integer = 200
    
    If Not mclsInsure.InitInsure(gcnOracle, TYPE_����) Then Exit Function
    Select Case lng���ܺ�
    Case C����_�ֹ������������
        '����������ֹ���µ����������ݶ�HIS�����ݵ�������ɲ���Ա��ǰ�û���������ţ��ڴ˴�¼�뼴����������������
        ���깤�߰� = ����_�ֹ������������
    End Select
End Function

Private Function ����_�ֹ������������() As Boolean
    Dim str������ As String
    Dim blnReturn As Boolean
    
    If gstrҽ���������� = "" Then
        MsgBox "׼����ȡҽ���������룬�����ϵͳ�����˿�", vbInformation, gstrSysName
CheckCard:
        initType
        blnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If blnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstrҽ���������� = gstrOutPara.out1
        gstrҽԺ���� = gstrOutPara.out2
    End If
        
    On Error GoTo errHandle
    str������ = InputBox("��¼������ţ�", "���ݾ����������������")
    If Trim(str������) = "" Then
        MsgBox "������Ϊ�գ��޷��������������ϣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���ýӿ�������
    initType
    blnReturn = fl_canrollback(gstrҽ����������, gstrҽԺ����, str������, gstrOutPara)
    TrimType
    If blnReturn = False Then
        MsgBox "�ж��Ƿ���Գ���ʱ��ҽ���˷���������Ϣ���˷Ѳ��ܼ�����" & Chr(13) & Chr(10) & gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If
    initType
    blnReturn = fl_rollbackcalc(gstrҽ����������, gstrҽԺ����, str������, "0", gstrOutPara)
    TrimType
    If blnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If

    ����_�ֹ������������ = True
    Exit Function
errHandle:
    MsgBox "��������[�ֹ������������]�����У�������Ϣ��" & Chr(13) & Chr(10) & Err.Description
End Function



