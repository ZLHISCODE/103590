VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmSentenceEdit 
   Caption         =   "�ʾ�ʾ���༭"
   ClientHeight    =   5535
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   8100
   Icon            =   "frmSentenceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   8100
   StartUpPosition =   1  '����������
   Begin MSComctlLib.TreeView tvw���� 
      Height          =   2430
      Left            =   1335
      TabIndex        =   17
      Tag             =   "1000"
      Top             =   1500
      Visible         =   0   'False
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   4286
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   12
      Text            =   "(��)"
      Top             =   1200
      Width           =   4980
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "���ൽ(&L)"
      Height          =   350
      Left            =   105
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1170
      Width           =   1215
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   855
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   795
      Width           =   3660
   End
   Begin VB.PictureBox pic���� 
      Height          =   3660
      Left            =   30
      ScaleHeight     =   3600
      ScaleWidth      =   7950
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1605
      Width           =   8010
      Begin zlRichEditor.Editor edt���� 
         Height          =   3060
         Left            =   75
         TabIndex        =   14
         Top             =   420
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   5398
         PaperHeight     =   11907
         PaperWidth      =   16840
         WithViewButtonas=   0   'False
         PaperKind       =   4
         ShowRuler       =   0   'False
         AuditMode       =   -1  'True
      End
      Begin XtremeCommandBars.CommandBars cbsThis 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VisualTheme     =   2
      End
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "����ʹ��(&3)"
      Height          =   180
      Index           =   2
      Left            =   4140
      TabIndex        =   7
      Top             =   525
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "����ͨ��(&2)"
      Height          =   180
      Index           =   1
      Left            =   2497
      TabIndex        =   6
      Top             =   525
      Width           =   1305
   End
   Begin VB.OptionButton opt��Χ 
      Caption         =   "ȫԺͨ��(&1)"
      Height          =   180
      Index           =   0
      Left            =   855
      TabIndex        =   5
      Top             =   525
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6705
      TabIndex        =   15
      Top             =   105
      Width           =   1215
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   3075
      TabIndex        =   3
      Top             =   105
      Width           =   3225
   End
   Begin VB.TextBox txt��� 
      Height          =   300
      Left            =   855
      TabIndex        =   1
      Top             =   105
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6705
      TabIndex        =   16
      Top             =   525
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgList 
      Bindings        =   "frmSentenceEdit.frx":058A
      Left            =   105
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceEdit.frx":059E
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceEdit.frx":0B38
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   855
      Width           =   630
   End
   Begin VB.Label lbl��Ա 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4560
      TabIndex        =   10
      Top             =   795
      Width           =   1740
   End
   Begin VB.Label lbl��Χ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ʹ��(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   4
      Top             =   525
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2400
      TabIndex        =   2
      Top             =   165
      Width           =   630
   End
   Begin VB.Label lbl��� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   0
      Top             =   165
      Width           =   630
   End
End
Attribute VB_Name = "frmSentenceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1���ϼ�����ͨ��������ShowMe�������������塢�༭�ʾ�ID,�༭״̬����Ϣ���ݽ��뱾����
'   2���༭״̬����Me.tag��ţ��ֱ�Ϊ"����"��"�޸�"�����ϼ�����ͨ��ShowMe����
'---------------------------------------------------
Private mlngClassId As Long                         '����ID
Private mlngWordId As Long                          '�ʾ�ID
Private mblnOK As Boolean                           '�Ƿ���ɱ༭�˳�

Private Elements As cEPRElements                    '�ֲ�����Ҫ�ؼ���
Private WithEvents mfrmInsElement As frmInsElement  '��������Ҫ�ش���
Attribute mfrmInsElement.VB_VarHelpID = -1
Private mlngHP As Long, blnSpaceEvent As Boolean    '��¼�Զ����ӿո��λ�ã�

Private blnActive As Boolean

'��ʱ����
Dim lngCount As Long

'-----------------------------------------------------
'����Ϊ�ⲿ��������
'-----------------------------------------------------
Public Function ShowMe(ByVal frmParent As Form, _
    ByVal blnAdd As Boolean, ByVal bytPower As Byte, ByVal lngClassId As Long, _
    Optional ByVal lngWordId As Long, Optional ByVal blnSaveAs As Boolean) As Long
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '������ bytPower-����Ȩ�ޣ�0-ȫԺ��1-���ң�2-���ˣ�
    '       lngClassId-����id
    '       lngWordId-��¼ID���޸�ʱ����
    '       blnSaveAs-�Ƿ����༭�����еġ����ʾ�ʾ��������
    '���أ�ȷ�������������޸ĵ�ID��ȡ������0
    '---------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    blnActive = False
    mlngClassId = lngClassId: mlngWordId = lngWordId
    If blnAdd Then
        Me.Tag = "����": mlngWordId = 0
    Else
        Me.Tag = "�޸�"
    End If
    
    '---------------------------------------------------
    '����������Ϣ
    Dim objNode As MSComctlLib.Node
    Err = 0: On Error GoTo ErrHand
    
    gstrSQL = "Select ID, �ϼ�id, ����, ����, ˵��" & vbNewLine & _
            "From �����ʾ����" & vbNewLine & _
            "Start With �ϼ�id Is Null" & vbNewLine & _
            "Connect By Prior ID = �ϼ�id" & vbNewLine & _
            "Order By Level, ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.tvw����.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvw����.Nodes.Add(, , "_" & !ID, !���� & "-" & !����, "close")
            Else
                Set objNode = Me.tvw����.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, !���� & "-" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            If !ID = lngClassId Then
                objNode.Selected = True
                Me.txt����.Tag = !ID: Me.txt����.Text = objNode.Text
            End If
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select Distinct d.Id, d.����, d.����, r.ȱʡ, r.��Աid, p.���� " _
            & "From ���ű� d, ������Ա r, ��Ա�� p, �ϻ���Ա�� u, ��������˵�� c " _
            & "Where d.Id = r.����id And r.��Աid = p.Id And p.Id = u.��Աid And u.�û��� = User And d.Id = c.����id And " _
            & "      c.�������� In ('�ٴ�', '���', '����', '����', '����', '����', 'Ӫ��', '���') And (p.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.����ʱ�� Is Null) " _
            & "Order By r.ȱʡ Desc,d.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.cbo����.Clear
        Do While Not .EOF
            Me.cbo����.AddItem !���� & "-" & !����
            Me.cbo����.ItemData(Me.cbo����.NewIndex) = !ID
            If !ȱʡ = 1 Then Me.cbo����.ListIndex = Me.cbo����.NewIndex
            Me.lbl��Ա.Tag = !��ԱID: Me.lbl��Ա.Caption = !����
            .MoveNext
        Loop
        If Me.cbo����.ListCount = 0 Then
            MsgBox "��Ŀǰ�������κ��ٴ�/���/����/����/����/����/Ӫ��/��첿�ţ����ܹ���ʾ�ʾ����", vbExclamation, gstrSysName
            ShowMe = 0: Unload Me: Exit Function
        ElseIf Me.cbo����.ListIndex = -1 Then
            Me.cbo����.ListIndex = 0
        End If
    End With
    
    '---------------------------------------------------
    '����������ȡ
    gstrSQL = "Select l.����id, l.���, l.����, l.ͨ�ü�, l.����id, d.����, d.���� As ����, l.��Աid, p.���� As ��Ա " _
            & "From �����ʾ�ʾ�� l, ���ű� d, ��Ա�� p " _
            & "Where l.����id = d.Id And l.��Աid = p.Id And l.id =[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngWordId)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt���.Text = !���
            Me.txt����.Text = !����
            Me.tvw����.Nodes("_" & !����id).Selected = True
            Me.txt����.Tag = !����id: Me.txt����.Text = Me.tvw����.SelectedItem.Text
            Me.opt��Χ(IIf(IsNull(!ͨ�ü�), 0, !ͨ�ü�)).Value = True
            If !��ԱID <> Me.lbl��Ա.Tag Then
                Me.lbl��Ա.Tag = !��ԱID: Me.lbl��Ա.Caption = !��Ա
                Me.cbo����.Clear
                Me.cbo����.AddItem !���� & "-" & !����
                Me.cbo����.ItemData(Me.cbo����.NewIndex) = !����ID
                Me.cbo����.ListIndex = Me.cbo����.NewIndex
                Me.cbo����.Enabled = False
            Else
                For lngCount = 0 To Me.cbo����.ListCount - 1
                    If Me.cbo����.ItemData(lngCount) = IIf(IsNull(!����ID), 0, !����ID) Then
                        Me.cbo����.ListIndex = lngCount: Exit For
                    End If
                Next
            End If
        End If
        Me.txt���.MaxLength = .Fields("���").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
    End With
    
    If InStr(1, gstrPrivsEpr, "ȫԺ�����ʾ�") <> 0 Then
        
    ElseIf InStr(1, gstrPrivsEpr, "���Ҳ����ʾ�") <> 0 Then
        Me.opt��Χ(0).Enabled = False
    ElseIf InStr(1, gstrPrivsEpr, "���˲����ʾ�") <> 0 Then
        Me.opt��Χ(0).Enabled = False: Me.opt��Χ(1).Enabled = False
    End If
    If Me.Tag = "����" Then Call zlDefaultCode
    
    '---------------------------------------------------
    '�ʾ����ݻָ�
    If blnAdd = False Then
        Call InsertPhrase(mlngWordId)
    ElseIf blnSaveAs Then
        Call InsertSelText(frmParent)
    End If
    
    '---------------------------------------------------
    Call InitMenu
    '---------------------------------------------------
    '��ʾ����
    Me.edt����.AuditMode = False
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
    If mblnOK Then
        ShowMe = mlngWordId
    Else
        ShowMe = 0
    End If
    Unload Me
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = 0
End Function

'-----------------------------------------------------
'����Ϊ�ڲ���������
'-----------------------------------------------------
Private Sub zlDefaultCode()
    '���ܣ�����Ĭ�ϵı��
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select LPad(Nvl(To_Number(Max(���)), 0) + 1, Nvl(Max(Length(���)), 5), '0') As ����" & vbNewLine & _
            "From �����ʾ�ʾ��" & vbNewLine & _
            "Where ����id = [1]"
    Err = 0: On Error Resume Next
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.txt����.Tag))
    Me.txt���.Text = rsTemp.Fields(0).Value
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.txt���.Text = ""
End Sub

Private Sub InitMenu()
    '���ܣ� �༭������������
    Dim rsTemp As New ADODB.Recordset
    Dim cbrControl As CommandBarControl, cbrCombox As CommandBarComboBox
    '---------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagStretched Or xtpFlagHideWrap
    With cbsThis.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlLabel, 0, "�ɵ���ʾ�ʾ��")
        Set cbrCombox = .Add(xtpControlComboBox, conMenu_Edit_Import, "ʾ���б�")
        cbrCombox.Width = 160: cbrCombox.DropDownWidth = 180: cbrCombox.DropDownListStyle = True
        gstrSQL = "Select Id, ���, ���� From �����ʾ�ʾ�� Where Id <> [1] And ����id = [2] Order By ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngWordId, mlngClassId)
        Do While Not rsTemp.EOF
            cbrCombox.AddItem rsTemp!��� & "-" & rsTemp!����
            cbrCombox.ItemData(rsTemp.AbsolutePosition) = rsTemp!ID
            rsTemp.MoveNext
        Loop
        cbrCombox.ListIndex = 1
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����Ҫ��(Ctrl+I)"): cbrControl.flags = xtpFlagRightAlign: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�Ҫ��(Ctrl+M)"): cbrControl.flags = xtpFlagRightAlign
    End With
    For Each cbrControl In cbsThis.ActiveMenuBar.Controls
        If cbrControl.Type = xtpControlButton Then cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("I"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("J"), ID_INSERT_AUTORECOGNISE
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        
        .Add FCONTROL, Asc("X"), ID_EDIT_CUT
        .Add FCONTROL, Asc("C"), ID_EDIT_COPY
        .Add FCONTROL, Asc("V"), ID_EDIT_PASTE
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F12, ID_INSERT_AUTORECOGNISE              '����ʶ��
    End With
End Sub

Private Sub ValidteRTF()
    '���RTF�е������ı��͹ؼ���
    Dim sType As String, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim i As Long, bFinded As Boolean
    i = 1
    
    Me.edt����.ForceEdit = True
    Do
        bFinded = FindNextAnyKey(edt����, i + 1, sType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            If sType <> "E" Then
                Me.edt����.Range(lKSS, lKEE).Font.Protected = False
                Me.edt����.Range(lKSS, lKEE).Font.ForeColor = tomAutoColor
                Me.edt����.Range(lKSS, lKEE).Font.BackColor = tomAutoColor
                Me.edt����.Range(lKSS, lKEE).Font.Hidden = False
                Me.edt����.Range(lKES, lKEE) = ""
                Me.edt����.Range(lKSS, lKSE) = ""
            Else
                i = lKEE
            End If
        Else
            i = i + 1
        End If
    Loop Until bFinded = False
    Me.edt����.ForceEdit = False
End Sub

Private Sub AppendElement(ByRef Ele As cEPRElement)
    '���Ҫ��
    Dim lngKey As Long, lngLen As Long
    lngLen = Len(Me.edt����.Text)
    Me.edt����.Range(lngLen, lngLen).Selected
    lngKey = Elements.AddExistNode(Ele, True)
End Sub

Private Sub AppendText(ByVal strText As String)
    '����ı�
    Dim lngKey As Long, lngLen As Long
    lngLen = Len(Me.edt����.Text)
    Me.edt����.ForceEdit = True
    Me.edt����.Range(lngLen, lngLen) = strText
    Me.edt����.ForceEdit = False
End Sub

'################################################################################################################
'## ���ܣ�  ��ʾ�Զ�ʶ������Ҫ�ػ����ֵ���Ŀ��ѡ����
'##
'## ������  strAuto     :IN     �����ѯ�ؼ���
'################################################################################################################
Private Sub ShowAutoRecSelector(ByVal strF As String)
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    bInKeys = IsBetweenAnyKeys(edt����, Me.edt����.SelStart + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then Exit Sub     '��֤���ܲ���ؼ����ڲ�
    If Me.edt����.Selection.Font.Protected Then Exit Sub

    Dim rs As New ADODB.Recordset
    Dim lLeft As Long, lTOp As Long
    
    '�������������Ӣ�����������������һЩ��
    gstrSQL = "select  ID,����,������ As ����,��λ,decode(�滻��,2,'�ֵ���Ŀ',1,'�滻��Ŀ','�ⲿ������') As ���� " & _
        "From ����������Ŀ " & _
        "Where ������ Like '%" & strF & "%' Or Ӣ���� Like '%" & UCase(strF) & "%' " & _
        "Order By ����"
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.EOF Then Exit Sub
    Dim pt As POINTAPI, arrPara As String, T As Variant, lngID As Long
    Dim f As New frmSelectChild
    
    pt.x = 0
    pt.y = 0
    ClientToScreen Me.edt����.OriginRTB.hwnd, pt
    '��ȡ��ʼλ������
    Me.edt����.Range(edt����.SelStart, Me.edt����.SelStart + 1).GetPoint cprGPStart + cprGPLeft + cprGPBottom, lLeft, lTOp

    arrPara = "0;830;2500;700;1000"
    strF = f.ShowSelectChild(Me, pt.x * Screen.TwipsPerPixelX + lLeft, pt.y * Screen.TwipsPerPixelY + lTOp, _
        5550, 3000, rs, arrPara)
    If strF = "" Then
        Exit Sub
    Else
        T = Split(strF, ";")
        lngID = T(0)
        rs.Close
        gstrSQL = "Select ID, ������, ����, ����, С��, ��λ, ��ʾ��, �滻��, ��ʼֵ, ��ֵ�� From ����������Ŀ Where ID =[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
        If Not rs.EOF Then
            '����Ԫ��
            Dim Ele As New cEPRElement, aryTemp() As String, lngKey As Long, lngCount As Long
            With Ele
                .Ҫ������ = NVL(rs("������"))
                .����Ҫ��ID = NVL(rs("ID"), 0)
                .Ҫ������ = NVL(rs("����"), 1)
                .Ҫ�س��� = NVL(rs("����"), 0)
                .Ҫ��С�� = NVL(rs("С��"), 0)
                .Ҫ�ص�λ = NVL(rs("��λ"))
                .Ҫ�ر�ʾ = IIf(NVL(rs("��ʾ��"), 0) = 4, 2, NVL(rs("��ʾ��"), 0))
                .�滻�� = NVL(rs("�滻��"), 0)      '0-�ⲿ������Ŀ��1-�滻��Ŀ��2-�ֵ���Ŀ
                .�����ı� = Trim(NVL(rs("��ʼֵ")))
                If .Ҫ������ = 0 Then
                    Select Case .Ҫ�ر�ʾ
                    Case 0, 1
                        If Trim(NVL(rs("��ֵ��"))) = "" Then
                            .Ҫ��ֵ�� = ""
                        Else
                            aryTemp = Split(NVL(rs("��ֵ��")), ";")
                            .Ҫ��ֵ�� = Val(aryTemp(0)) & ";" & Val(aryTemp(1))
                        End If
                    Case 2
                        aryTemp = Split(NVL(rs("��ֵ��")), ";")
                        For lngCount = 0 To UBound(aryTemp)
                            aryTemp(lngCount) = Val(aryTemp(lngCount))
                        Next
                        .Ҫ��ֵ�� = Join(aryTemp(0), ";")
                    Case Else
                        .Ҫ��ֵ�� = ""
                    End Select
                Else
                    Select Case .Ҫ�ر�ʾ
                    Case 2, 3
                        .Ҫ��ֵ�� = NVL(rs("��ֵ��"))
                    Case Else
                        .Ҫ��ֵ�� = ""
                    End Select
                End If
                .������̬ = IIf(.Ҫ�ر�ʾ = 2 Or .Ҫ�ر�ʾ = 3, 1, 0) '0-�ı� 1-���� 2-��ѡ 3-��ѡ   ���Ϊ��ѡ����ѡ��������Ĭ��ֵΪչ����Ŀ   0-����;1-չ��
            End With
            lngKey = Elements.AddExistNode(Ele)
            
            '��������Ҫ�ص��༭����
            Dim blnForce As Boolean
            blnForce = Me.edt����.ForceEdit
            Me.edt����.ForceEdit = True
            Me.edt����.SelText = ""
            Elements("K" & lngKey).InsertIntoEditor Me.edt����, , True
            Me.edt����.ForceEdit = blnForce
        End If
    End If
End Sub

'################################################################################################################
'## ���ܣ���ָ���ʾ�ʾ������༭�������ڴʾ���޸����ݻָ��������ʾ����
'################################################################################################################

Private Sub InsertPhrase(ByVal lngImpId As Long)
    Dim rsTemp As New ADODB.Recordset
    
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim lngKey As Long, lngStart As Long, lngLen As Long, strTmp As String
    
    bInKeys = IsBetweenAnyKeys(edt����, Me.edt����.SelStart + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then Exit Sub
    
    gstrSQL = "Select �ʾ�id, ���д���, ��������, �����ı�, ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, Ҫ��ֵ��, ������̬, ��������" & vbNewLine & _
                "From �����ʾ����" & vbNewLine & _
                "Where �ʾ�id = [1]" & vbNewLine & _
                "Order By ���д���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngImpId)
    With Me.edt����
        .Freeze
        .ForceEdit = True
        Do While Not rsTemp.EOF
            Select Case rsTemp("��������")
            Case 0 '��������
                '�ָ�RTF����
                lngStart = .SelStart
                strTmp = NVL(rsTemp("�����ı�"))
                lngLen = Len(strTmp)

                .Range(lngStart, lngStart) = strTmp
                .Range(lngStart, lngStart + lngLen).Font.Protected = False
                .Range(lngStart, lngStart + lngLen).Font.Hidden = False
                .Range(lngStart + lngLen, lngStart + lngLen).Selected
            Case 1, 2 '1-��ʱ����Ҫ��,2-�̶�����Ҫ��
                lngStart = .SelStart
                
                lngKey = Elements.Add
                Elements("K" & lngKey).ID = mlngWordId
                Elements("K" & lngKey).�����ı� = NVL(rsTemp("�����ı�"))
                Elements("K" & lngKey).Ҫ������ = NVL(rsTemp("Ҫ������"))
                Elements("K" & lngKey).����Ҫ��ID = NVL(rsTemp("����Ҫ��ID"), 0)
                Elements("K" & lngKey).�滻�� = NVL(rsTemp("�滻��"), 0)
                Elements("K" & lngKey).Ҫ������ = NVL(rsTemp("Ҫ������"), 0)
                Elements("K" & lngKey).Ҫ�س��� = NVL(rsTemp("Ҫ�س���"), 0)
                Elements("K" & lngKey).Ҫ��С�� = NVL(rsTemp("Ҫ��С��"), 0)
                Elements("K" & lngKey).Ҫ�ص�λ = NVL(rsTemp("Ҫ�ص�λ"))
                Elements("K" & lngKey).Ҫ�ر�ʾ = NVL(rsTemp("Ҫ�ر�ʾ"), 0)
                Elements("K" & lngKey).Ҫ��ֵ�� = NVL(rsTemp("Ҫ��ֵ��"))
                Elements("K" & lngKey).������̬ = NVL(rsTemp("������̬"), 0)
                Elements("K" & lngKey).�Ƿ��� = False
                Elements("K" & lngKey).�������� = NVL(rsTemp!��������)
                Elements("K" & lngKey).InsertIntoEditor Me.edt����, lngStart, , True
            
            End Select
            rsTemp.MoveNext
        Loop
        lngStart = .SelStart
        .ForceEdit = False
        
        'ȥ��ĩβ�Ļس����з�
        lngLen = Len(Me.edt����.Text)
        If lngLen > 0 Then
            If (Me.edt����.Range(lngLen - 2, lngLen) = vbCrLf Or (Asc(Me.edt����.Range(lngLen - 1, lngLen)) = 13 And Asc(Me.edt����.Range(lngLen - 2, lngLen - 1)) = 10)) And Me.edt����.Range(lngLen - 2, lngLen).Font.Protected = False Then
                Me.edt����.Range(lngLen - 2, lngLen) = ""
            End If
        End If
        
        .Range(lngStart, lngStart).Selected
        .Modified = False
        .UnFreeze
    End With
End Sub

'################################################################################################################
'## ���ܣ����ϼ������ѡ������ָ����ʽ�ı����뵽�༭����
'################################################################################################################
Private Sub InsertSelText(ByVal frmEdit As Form)
    Dim sType As String, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, bFinded As Boolean
    Dim lngKey As Long, bBeteenKeys As Boolean
    
    '��չ��֤������Ҫ��ѡ��
    lS = frmEdit.Editor1.Selection.StartPos
    lE = frmEdit.Editor1.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(frmEdit.Editor1, lS + 1, sType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(frmEdit.Editor1, lE + 1, sType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    
    '�ȸ�ֵRTF
    Me.edt����.NewDoc
    Me.edt����.ForceEdit = True
    Me.edt����.TOM.TextDocument.Selection.FormattedText = frmEdit.Editor1.TOM.TextDocument.Range(lS, lE).FormattedText
    SetCommonStyle Me.edt����, "����", 0, Len(Me.edt����.Text), True
    
    '����Ҫ��
    Set Elements = New cEPRElements
    For i = lS To lE
        bFinded = FindNextAnyKey(frmEdit.Editor1, i + 1, sType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded = False Then Exit For    '�������κ�Ԫ�أ���ô�˳�ѭ��
        If Not (lKSS < lE) Then Exit For    '������Χ���˳�ѭ��
        '��Χ�ڴ��ڹؼ���
        If sType = "E" Then
            '�����Ҫ�أ���ô������������
            Call AppendElement(frmEdit.Document.Elements("K" & lKey).Clone(True))
        End If
        i = lKEE - 1
    Next
    Call ValidteRTF
End Sub

'################################################################################################################
'## ���ܣ�  ��ȡ���ı������ݿ��SQL���
'##
'## ������
'##         ArraySQL()      :IN/OUT��   SQL����
'##         strIn           :IN��       ��Ҫ������ַ���
'##         lng���         :IN��       ���
'##         bln�Ƿ���     :IN��       �Ƿ���
'##
'## ˵����  ���ȴ���4000���ַ��������д洢����ŵ���֮��
'################################################################################################################
Private Function GetPlainTextSaveSQL(ByRef ArraySQL() As String, _
    ByVal strIn As String, ByRef lng��� As Long) As Boolean
    
    Dim lngLen As Long, strSub As String, i As Long, lngID As Long
    Dim lngCount As Long, lID As Long
    strIn = Replace(strIn, "'", "' || chr(39) || '")
    strIn = Replace(strIn, vbCrLf, "' || chr(13) || chr(10) || '")  '����strIn�ǲ�������vbCrlf�ġ�
    strIn = Replace(strIn, "��", " ") '����ȫ�ǿո�
    lngLen = Len(strIn)
    
    '����4000Ϊ��ֶδ洢��
    i = 0
    Do While (i * 2000 + 1 <= lngLen)
        lngCount = UBound(ArraySQL) + 1
        ReDim Preserve ArraySQL(1 To lngCount) As String

        strSub = Mid(strIn, i * 2000 + 1, 2000)

        gstrSQL = "Zl_�����ʾ����_Insert(" & mlngWordId & "," & lng��� & ",0,'" & strSub & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)"
        
        ArraySQL(lngCount) = gstrSQL
       
        lng��� = lng��� + 1
        i = i + 1
    Loop
    GetPlainTextSaveSQL = True
End Function

'################################################################################################################
'## ���ܣ�  ���ݵĸ��Ʋ����������ı���Ҫ�أ�
'################################################################################################################
Private Sub ExecCopy()
    If Me.edt����.ReadOnly Then Exit Sub
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, bFinded As Boolean, lngLen As Long, lngSum As Long
    
    '��չ��ʼλ�ú���ֹλ�ã�ʹ�������������Ҫ�ض���
    lS = Me.edt����.Selection.StartPos
    lE = Me.edt����.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(edt����, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(edt����, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    
    '�ȿ���RTF����
    gfrmPublic.edtPublic.NewDoc
    gfrmPublic.edtPublic.ForceEdit = True
    gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText = Me.edt����.TOM.TextDocument.Range(lS, lE).FormattedText
    '����Ҫ�أ���������Ԫ�أ�ͼƬ����ϡ����ȣ����ؼ���ҲҪ������ȥ����֤�����ݵ����عؼ���Keyֵһ�£�
    Set gfrmPublic.Elements = New cEPRElements
    lngSum = 0
    For i = lS To lE
        bFinded = FindNextAnyKey(edt����, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                '��Χ�ڴ��ڹؼ���
                If sKeyType = "E" Then
                    '�����Ҫ�أ���ô������������
                    gfrmPublic.Elements.AddExistNode Elements("K" & lKey).Clone(True), True
                Else
                    '���������Ԫ�أ������֮����gfrmPublic.edtPublic�����������¼��ǰλ�ã���
                    gfrmPublic.edtPublic.Range(lKSS - lS - lngSum, lKEE - lS - lngSum) = ""
                    lngSum = lngSum + lKEE - lKSS   '��¼ɾ�����ݵ��ܳ���
                End If
            Else
                '���򣬳�����Χ���˳�ѭ��
                Exit For
            End If
            i = lKEE - 1
        Else
            '�������κ�Ԫ�أ���ô�˳�ѭ��
            Exit For
        End If
    Next
    Clipboard.Clear
End Sub

'################################################################################################################
'## ���ܣ�  ���ݵļ��в����������ı���Ҫ�أ�
'################################################################################################################
Private Sub ExecCut()
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, bFinded As Boolean, lngNum As Long, lngSum As Long
    
    '��չ��ʼλ�ú���ֹλ�ã�ʹ�������������Ҫ�ض���
    lS = Me.edt����.Selection.StartPos
    lE = Me.edt����.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(edt����, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(edt����, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    
    '�ȿ���RTF����
    gfrmPublic.edtPublic.NewDoc
    gfrmPublic.edtPublic.ForceEdit = True
    gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText = Me.edt����.TOM.TextDocument.Range(lS, lE).FormattedText
    '����Ҫ�أ���������Ԫ�أ�ͼƬ����ϡ����ȣ����ؼ���ҲҪ������ȥ����֤�����ݵ����عؼ���Keyֵһ�£�
    Set gfrmPublic.Elements = New cEPRElements
    lngSum = 0
    For i = lS To lE
        bFinded = FindNextAnyKey(edt����, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                '��Χ�ڴ��ڹؼ���
                If sKeyType = "E" Then
                    '�����Ҫ�أ���ô������������
                    gfrmPublic.Elements.AddExistNode Elements("K" & lKey), True
                Else
                    '���������Ԫ�أ������֮����gfrmPublic.edtPublic�����������¼��ǰλ�ã���
                    gfrmPublic.edtPublic.Range(lKSS - lS - lngSum, lKEE - lS - lngSum) = ""
                    lngSum = lngSum + lKEE - lKSS   '��¼ɾ�����ݵ��ܳ���
                End If
            Else
                '���򣬳�����Χ���˳�ѭ��
                Exit For
            End If
            i = lKEE - 1
        Else
            '�������κ�Ԫ�أ���ô�˳�ѭ��
            Exit For
        End If
    Next
    
    'ɾ��ѡ������
    Dim bForce As Boolean, COLOR As OLE_COLOR, bProtect1 As Boolean, bProtect2 As Boolean
    bForce = Me.edt����.ForceEdit
    Me.edt����.ForceEdit = True
    Me.edt����.Range(lS, lE) = ""
    Me.edt����.ForceEdit = bForce
    Clipboard.Clear
End Sub

'################################################################################################################
'## ���ܣ�  ���ݵ�ɾ������
'################################################################################################################
Private Sub ExecDelete()
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, j As Long, bFinded As Boolean, lngNum As Long, lngSum As Long
    
    '��չ��ʼλ�ú���ֹλ�ã�ʹ�������������Ҫ�ض���
    lS = Me.edt����.Selection.StartPos
    lE = Me.edt����.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(edt����, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(edt����, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    
    'ɾ��ѡ������
    Dim bForce As Boolean, COLOR As OLE_COLOR, bProtect1 As Boolean, bProtect2 As Boolean
    bForce = Me.edt����.ForceEdit
    Me.edt����.ForceEdit = True
    
    If Me.edt����.SelLength > 0 Then
        'ѡ�����ݷǿ�
        '��չ��ʼλ�ú���ֹλ�ã�ʹ�������������Ҫ�ض���
        lS = Me.edt����.Selection.StartPos
        lE = Me.edt����.Selection.EndPos
        bBeteenKeys = IsBetweenAnyKeys(edt����, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lS = lKSS
        bBeteenKeys = IsBetweenAnyKeys(edt����, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lE = lKEE
        
        Me.edt����.Freeze
        '���޶�ģʽ�����������Ҫ�ء�ͼƬ�������ϣ�����ɾ�����
        lngSum = 0
        For i = lS To lE - 1
            bFinded = FindNextAnyKey(edt����, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bFinded Then
                If lKSS < lE Then   '��Χ�ڴ��ڹؼ���
                    '1���ȴ���ǰ�������
                    Me.edt����.Range(i, lKSS) = ""
                    lngNum = lKSS - i
                    lE = lE - lngNum
                    lngSum = lngSum + lngNum
                    i = lKSS - lngNum - 1
                    '2���������һ��Ҫ�ء�ͼƬ��������
                    Select Case sKeyType
                    Case "E"    'Ҫ��
                        If Elements("K" & lKey).�������� = False Then
                            Me.edt����.Range(lKSS - lngNum, lKEE - lngNum) = ""
                            Elements.Remove "K" & lKey
                            lngSum = lngSum + (lKEE - lKSS)
                            lE = lE - (lKEE - lKSS)
                        Else
                            i = lKEE - lngNum - 1
                        End If
                    Case Else
                       '���������Ԫ�أ��򲻴���
                       i = lKEE - lngNum - 1
                    End Select
                Else
                    '���򣬳�����Χ���˳�ѭ��
                    Exit For
                End If
            Else
                '�������κ�Ԫ�أ���ô�˳�ѭ��
                Exit For
            End If
        Next
        If i < lE Then
            Me.edt����.Range(i, lE) = ""
            lngNum = lE - i
        End If
        Me.edt����.UnFreeze
        Me.edt����.SelLength = 0
        Me.edt����.Range(lE - lngNum, lE - lngNum).Selected
        Clipboard.Clear
    Else
        'û��ѡ���ı�
        bBeteenKeys = IsBetweenAnyKeys(edt����, Me.edt����.SelStart + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then
            'ɾ����������Ҫ��
            Select Case sKeyType
            Case "E"
                Elements.Remove "K" & lKey
            Case Else
                GoTo LL
            End Select
            Me.edt����.Range(lKSS, lKEE) = ""
            If Me.edt����.Range(lKSS - 2, lKSS) = vbCrLf And Me.edt����.Range(lKSS - 2, lKSS).Font.Protected Then
                Me.edt����.Range(lKSS - 2, lKSS) = ""
                Me.edt����.Range(lKSS - 2, lKSS - 2).Font.Protected = False
            Else
                Me.edt����.Range(lKSS, lKSS).Font.Protected = False
            End If
        Else
            'ɾ���ı�
            i = Me.edt����.SelStart
            j = Len(edt����.Text)
            
            If Me.edt����.Range(i, i + 1).Font.Protected = False And (edt����.Range(i + 1, i + 2).Font.Protected = True Or i = j - 1) Then
                Me.edt����.Range(i, i + 1) = ""
            ElseIf Me.edt����.Range(i - 1, i).Font.Protected = True And Me.edt����.Range(i, i + 1).Font.Protected = False Then
                If Me.edt����.Range(i, i + 2) = vbCrLf And Me.edt����.Range(i, i + 2).Font.Protected = False Then
                    Me.edt����.Range(i, i + 2) = ""
                    Me.edt����.Range(i, i).Font.Protected = False
                Else
                    Me.edt����.Delete
                End If
            ElseIf Me.edt����.Range(i, i + 2) = vbCrLf And Me.edt����.Range(i, i + 2).Font.Protected = False Then
                Me.edt����.Range(i, i + 2) = ""
                Me.edt����.Range(i, i).Font.Protected = False
            ElseIf Me.edt����.Range(i, i).Font.Protected = False And Me.edt����.Range(i, i + 1).Font.Protected = False Then
                Me.edt����.Delete
            ElseIf Me.edt����.Range(i, i + 2) = vbCrLf And Me.edt����.Range(i, i + 2).Font.Protected Then
                Me.edt����.Range(i + 2, i + 2).Selected
            Else
                Me.edt����.Range(i + 1, i + 1).Selected
            End If
        End If
    End If
LL:
    Me.edt����.ForceEdit = bForce
End Sub

'################################################################################################################
'## ���ܣ�  ���ݵ�ճ������������Ҫ�عؼ��֣�����ɾ����Ҫ��Ҫ����Ϊ�������޶��ı�Ҳͳһ��Ϊ�����ı���
'################################################################################################################
Private Sub ExecPaste(ByRef edtThis As Object)
    If edtThis.ReadOnly Then Exit Sub
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim i As Long, bForce As Boolean, bFinded As Boolean, strTmp As String, lS As Long, lE As Long, lngLen As Long
    Dim ParaFmt As New cParaFormat
    bBeteenKeys = IsBetweenAnyKeys(edtThis, edtThis.SelStart + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then Exit Sub    '������ճ����Ԫ���ڲ�
    
    If edtThis.Selection.Font.ForeColor = tomUndefined Or edtThis.Selection.Font.Protected Then Exit Sub
    
    '���������Ϊ�գ���ô��ճ���ڲ�����
    Dim strClipboard As String
    strClipboard = Clipboard.GetText
    If Len(Trim(strClipboard)) > 0 Then
        'ճ������������
        lS = edtThis.Selection.StartPos
        lE = lS + Len(strClipboard)
        edtThis.ForceEdit = True
        edtThis.Range(lS, edtThis.Selection.EndPos).Text = strClipboard
        edtThis.Range(lS, lE).Font.Strikethrough = False
        edtThis.Range(lS, lE).Font.Protected = False
        edtThis.Range(lS, lE).Font.ForeColor = tomAutoColor
        edtThis.ForceEdit = False
        edtThis.Range(lE, lE).Selected
        Exit Sub
    End If
    
    '�������ؼ���
    gfrmPublic.edtPublic.ForceEdit = True
    For i = 1 To gfrmPublic.Elements.Count
        '����Ҫ��
        lKey = Elements.AddExistNode(gfrmPublic.Elements(i).Clone, False)
        Elements("K" & lKey).��ʼ�� = 1
        Elements("K" & lKey).��ֹ�� = 0     'ȥ����ֹ��
        Elements("K" & lKey).�������� = False
        Elements("K" & lKey).ID = 0
        '�����ؼ���
        bFinded = FindKey(gfrmPublic.edtPublic, "E", gfrmPublic.Elements(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            strTmp = Format(lKey, "00000000") & "," & IIf(Elements("K" & lKey).��������, 1, 0) & ",0)"
            gfrmPublic.edtPublic.Range(lKSS, lKSE) = "ES(" & strTmp
            gfrmPublic.edtPublic.Range(lKES, lKEE) = "EE(" & strTmp
            gfrmPublic.Elements(i).Key = lKey '�����ı���ͬʱ������Key
        End If
    Next
    
    '����RTF���ݣ����ǰ��ɫ��ɾ����
    bForce = edtThis.ForceEdit
    edtThis.Freeze
    edtThis.ForceEdit = True
    
    lS = 0: lE = Len(gfrmPublic.edtPublic.Text)
    For i = lS To lE - 1
        If gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And gfrmPublic.edtPublic.Range(i, i + 1).Font.Protected Then
            '�����ı�ȥ������
            gfrmPublic.edtPublic.Range(i, i + 1).Font.Protected = False
        End If
        gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = tomAutoColor
    Next
    
    gfrmPublic.edtPublic.SelectAll
    gfrmPublic.edtPublic.Selection.Font.Strikethrough = False
    lS = edtThis.SelStart

    lngLen = Len(gfrmPublic.edtPublic.Text)
    If lngLen > 0 Then
        edtThis.TOM.TextDocument.Selection.FormattedText = gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText
        'ȥ��ĩβ�Ļس����з�
        If edtThis.Range(lS + lngLen, lS + lngLen + 2) = vbCrLf And edtThis.Range(lS + lngLen, lS + lngLen + 2).Font.Protected = False Then
            edtThis.Range(lS + lngLen, lS + lngLen + 2) = ""
        End If
        edtThis.Range(lS + lngLen, lS + lngLen).Selected
    End If
    lngLen = Len(edt����.Text)
    Me.edt����.Range(0, lngLen).Para.SetIndents 0, 0, 0
    Me.edt����.Range(0, lngLen).Para.SetLineSpacing cprLSSignle, 1
    Me.edt����.Range(0, lngLen).Para.ListType = cprLTNone
    Me.edt����.Range(0, lngLen).Font.Size = 10.5
    Me.edt����.Range(0, lngLen).Font.Name = "����"
    
    Me.edt����.ForceEdit = bForce
    Me.edt����.UnFreeze
End Sub


'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
'    On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        bInKeys = IsBetweenAnyKeys(Me.edt����, Me.edt����.SelStart + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
        If bInKeys Then Exit Sub
        If bInKeys = False Then mfrmInsElement.ShowMe Me, , True, False, True
    Case conMenu_Edit_Modify
        bInKeys = IsBetweenAnyKeys(Me.edt����, Me.edt����.SelStart + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
        If bInKeys Then
            mfrmInsElement.Tag = lKey
            mfrmInsElement.ShowMe Me, Elements("K" & lKey), True, False, True
        End If
    Case ID_EDIT_CUT
        ExecCut
    Case ID_EDIT_COPY
        ExecCopy
    Case ID_EDIT_PASTE
        ExecPaste Me.edt����
    Case ID_INSERT_AUTORECOGNISE                 '����ʶ��
        '�Զ�ʶ������Ҫ�ػ����ֵ���Ŀ
        Dim strAuto As String
        strAuto = Trim(Me.edt����.SelText)
        If strAuto = "" Then Exit Sub
        If Len(strAuto) > 100 Then strAuto = Left(strAuto, 100)
        ShowAutoRecSelector strAuto
    Case conMenu_Edit_Delete
        Call ExecDelete
    Case conMenu_Edit_Import
        If Control.Type <> xtpControlComboBox Then Exit Sub
        Call InsertPhrase(Control.ItemData(Control.ListIndex))
    End Select
    Me.edt����.SetFocus
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    Me.edt����.Move lngScaleLeft, lngScaleTop, lngScaleRight - lngScaleLeft, lngScaleBottom - lngScaleTop
    Me.edt����.PaperWidth = lngScaleRight - lngScaleLeft
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    If Me.edt����.Modified Then
        If MsgBox("�ʾ�ʾ�������Ѿ����޸ģ��Ƿ񱣴棿", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            Call cmdOK_Click
            Exit Sub
        End If
    End If
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim strText As String, ArraySQL() As String, i As Long, lngCount As Long, blnTran As Boolean
    
    If Trim(Me.txt���.Text) = "" Then MsgBox "�������ţ�", vbInformation, gstrSysName: Me.txt���.SetFocus: Exit Sub
'    If Len(Me.txt���.Text) < Me.txt���.MaxLength Then MsgBox "��ų��Ȳ��㣡", vbInformation, gstrSysName: Me.txt���.SetFocus: Exit Sub
    If Trim(Me.txt����.Text) = "" Then MsgBox "���������ƣ�", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    End If
    If Me.cbo����.ListIndex = -1 Then MsgBox "��������ң�", vbInformation, gstrSysName: Me.cbo����.SetFocus: Exit Sub
    
    '���ݱ���
    If Me.Tag = "����" Then
        mlngWordId = zlDatabase.GetNextId("�����ʾ�ʾ��")
        gstrSQL = mlngWordId & "," & Val(Me.txt����.Tag) & ",'" & Trim(Me.txt���.Text) & "','" & Trim(Me.txt����.Text) & "'"
        If Me.opt��Χ(0).Value Then
            gstrSQL = gstrSQL & ",0"
        ElseIf Me.opt��Χ(1).Value Then
            gstrSQL = gstrSQL & ",1"
        Else
            gstrSQL = gstrSQL & ",2"
        End If
        gstrSQL = gstrSQL & "," & Me.cbo����.ItemData(Me.cbo����.ListIndex) & "," & Me.lbl��Ա.Tag
        gstrSQL = "Zl_�����ʾ�ʾ��_Edit(1," & gstrSQL & ")"
    Else
        gstrSQL = mlngWordId & "," & Val(Me.txt����.Tag) & ",'" & Trim(Me.txt���.Text) & "','" & Trim(Me.txt����.Text) & "'"
        If Me.opt��Χ(0).Value Then
            gstrSQL = gstrSQL & ",0"
        ElseIf Me.opt��Χ(1).Value Then
            gstrSQL = gstrSQL & ",1"
        Else
            gstrSQL = gstrSQL & ",2"
        End If
        gstrSQL = gstrSQL & "," & Me.cbo����.ItemData(Me.cbo����.ListIndex)
        gstrSQL = "Zl_�����ʾ�ʾ��_Edit(2," & gstrSQL & ")"
    End If
    
    '��ȡSQL�������
    ReDim ArraySQL(1 To 2) As String
    ArraySQL(1) = gstrSQL
    
    'ǰ�ڴ���
    ArraySQL(2) = "Zl_�����ʾ����_Beforesave(" & mlngWordId & ")"
    
    '��ȡ����SQL����
    Call GetSaveSQL(ArraySQL)
    
    '���ڴ���
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String
    gstrSQL = "Zl_�����ʾ����_Aftersave(" & mlngWordId & ")"
    ArraySQL(lngCount) = gstrSQL
    
    'ִ�б������
    Err = 0: On Error GoTo ErrHand
    gcnOracle.BeginTrans
    blnTran = True
    For i = 1 To UBound(ArraySQL)
        gstrSQL = ArraySQL(i)
        Call zlDatabase.ExecuteProcedure(gstrSQL, "cEPRDocument")
    Next
    gcnOracle.CommitTrans
    blnTran = False
    mblnOK = True: Me.Hide
    Exit Sub

ErrHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetSaveSQL(ByRef ArraySQL() As String)
    '��ȡ����SQL���
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lngEnd As Long, M As Long, N As Long, p As Long
    Dim lng��� As Long, strText As String
    Dim lngCount As Long
    
    lng��� = 1     '����CRLF���ֶ�
    strText = Me.edt����.Text
    p = 0
    lngEnd = Len(Me.edt����.Text)
    Do While p < lngEnd
        '��ȡ�ؼ���λ�� M
        bFinded = FindNextAnyKey(Me.edt����, p + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            M = lKSS
        Else
            M = lngEnd
        End If
        
        '��ȡvbCrlfλ�� N
        N = InStr(p + 1, strText, vbCrLf, vbTextCompare)
        If N > 0 Then
            N = N - 1
        Else
            N = lngEnd
        End If
        
        If M < N Then
            '�����ı�
            Call GetPlainTextSaveSQL(ArraySQL, Me.edt����.Range(p, M), lng���)    '����Զ���1
            '�������
            If bFinded Then
                Select Case sKeyType
                Case "E"
                    Elements("K" & lKey).������� = lng���
                    p = lKEE    '������ǰλ��
                    With Elements("K" & lKey)
                        lngCount = UBound(ArraySQL) + 1
                        ReDim Preserve ArraySQL(1 To lngCount) As String
                        gstrSQL = "Zl_�����ʾ����_Insert(" & mlngWordId & "," & lng��� & ",1,'" & .�����ı� & "','" & .Ҫ������ & "'," & _
                            IIf(.����Ҫ��ID = 0, "NULL", .����Ҫ��ID) & "," & .�滻�� & "," & .Ҫ������ & "," & .Ҫ�س��� & "," & .Ҫ��С�� & ",'" & .Ҫ�ص�λ & "'," & _
                            .Ҫ�ر�ʾ & ",'" & .Ҫ��ֵ�� & "'," & .������̬ & ",'" & .�������� & "')"
                        ArraySQL(lngCount) = gstrSQL
                    End With
                    lng��� = lng��� + 1
                End Select
            Else
                p = M
            End If
        Else
            If Me.edt����.Range(N, N + 2) = vbCrLf And Me.edt����.Range(N, N + 2).Font.Protected = True Then
                '�ûس�������һ������ͼƬ���߱��
                If p < N Then Call GetPlainTextSaveSQL(ArraySQL, Me.edt����.Range(p, N), lng���)                 '����Զ���1
            Else
                '�����ı�
                Call GetPlainTextSaveSQL(ArraySQL, Me.edt����.Range(p, IIf(N >= lngEnd, N, N + 2)), lng���) '����Զ���1
            End If
            p = N + 2
        End If
    Loop
End Sub

Private Sub cmd����_Click()
    With Me.tvw����
        .Left = Me.txt����.Left: .Width = Me.txt����.Width
        .Top = Me.txt����.Top + Me.txt����.Height: .Height = Me.pic����.Height
        .ZOrder 0: .Visible = True: .SetFocus
    End With
End Sub

Private Sub edt����_Change(ViewMode As zlRichEditor.ViewModeEnum)
    If mlngHP > 0 Then
        If blnSpaceEvent Then
            blnSpaceEvent = False
            Exit Sub
        Else
            '�ָ��ո�
            If Me.edt����.Range(mlngHP - 1, mlngHP).Font.Hidden And _
                Me.edt����.Range(mlngHP, mlngHP + 1).Font.Hidden = False And _
                Me.edt����.Range(mlngHP, mlngHP + 1) = " " Then
                
                Dim blnForce As Boolean
                blnForce = Me.edt����.ForceEdit
                Me.edt����.ForceEdit = True
                Me.edt����.Range(mlngHP, mlngHP + 1) = ""
                Me.edt����.Range(mlngHP, mlngHP).Font.Protected = False
                Me.edt����.Range(mlngHP, mlngHP).Font.Hidden = False
                Me.edt����.ForceEdit = blnForce
            End If
            mlngHP = 0
            blnSpaceEvent = False
        End If
    End If
End Sub

Private Sub edt����_KeyDown(ViewMode As zlRichEditor.ViewModeEnum, KeyCode As Integer, Shift As Integer)
    If Me.edt����.SelLength > 0 Then Exit Sub
    If Shift <> 0 Then Exit Sub
    Select Case KeyCode
    Case 0, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
        vbKeyDelete, vbKeyBack, vbKeyTab, vbKeyInsert, vbKeyPageDown, vbKeyPageUp, _
        vbKeyPause, vbKeyPrint, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital, _
        vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, vbKeyF10, vbKeyF11, vbKeyF12
        
        DoEvents
        Exit Sub
    End Select
    
    '���������عؼ��ֺ��棬���Զ���һ���ո񣨷Ǳ������������ԣ�
    Dim i As Long, blnForce As Boolean
    With Me.edt����
        blnForce = .ForceEdit
        i = .SelStart
    
LL1:
        If .Range(i - 1, i).Font.Hidden And _
            .Range(i, i + 1).Font.Hidden = False And _
            .Range(i, i + 1).Font.Protected = False Then
            'A���⣺�������ı���|��ͨ�ı�
            
            mlngHP = i
            .ForceEdit = True
            .Range(i, i).Font.Protected = False
            .Range(i, i).Font.Hidden = False
            blnSpaceEvent = True
            .Range(i, i) = " "
            .Range(i + 1, i + 1).Selected
            .ForceEdit = blnForce
        Else
            If .Range(i - 1, i).Font.Hidden And _
                .Range(i, i + 1).Font.Hidden = False And _
                .Range(i, i + 1).Font.Protected Then
                'B����1����ͨ�ı��������ı���|�������ı����������ı�����ͨ�ı�
                i = i - 16
                If .Range(i - 1, i + 3) Like ")?S(" And _
                    .Range(i - 1, i + 3).Font.Hidden = True Then
                    'C���⣺�������ı����������ı����������ı���|�������ı����������ı����������ı���
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "��")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    .Range(i + 1, i + 1).Selected
                    .ForceEdit = blnForce
                ElseIf .Range(i + 1, i + 3) = "E(" And .Range(i, i + 3).Font.Protected And _
                    .Range(i + 16, i + 18) = vbCrLf And .Range(i + 16, i + 18).Font.Protected Then
                    'D���⣺��ٺ����ͼƬ����֮��û������ʱ���޷�������������
                    i = i + 16
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "��")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    If (.Range(i - 16, i - 14) <> "EE") Then
                        .Range(i, i + 1).Selected
                    Else
                        .Range(i + 1, i + 1).Selected
                    End If
                    .ForceEdit = blnForce
                Else
                    .Range(i, i).Selected
                End If
            ElseIf .Range(i - 1, i).Font.Hidden = False And _
                .Range(i - 1, i).Font.Protected And _
                .Range(i, i + 1).Font.Hidden Then
                'B����2����ͨ�ı��������ı����������ı���|�������ı�����ͨ�ı�
                i = i + 16
                If .Range(i - 1, i + 3) Like ")?S(" And _
                    .Range(i - 1, i + 3).Font.Hidden = True Then
                    'C���⣺�������ı����������ı����������ı���|�������ı����������ı����������ı���
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "��")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    .Range(i + 1, i + 1).Selected
                    .ForceEdit = blnForce
                Else
                    GoTo LL1
                End If
            ElseIf .Range(i - 1, i).Font.Hidden = False And .Range(i, i + 2) = vbCrLf And .Range(i, i + 2).Font.Protected Then
                mlngHP = -1
                .ForceEdit = True
                .Range(i, i).Font.Protected = False
                .Range(i, i).Font.Hidden = False
                .Range(i - 1, i).Font.ForeColor = vbBlack
                blnSpaceEvent = True
                .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "��")
                .Range(i, i + 1).Font.Protected = False
                .Range(i, i + 1).Font.Hidden = False
                .Range(i, i + 1).Font.ForeColor = vbBlack
                If (.Range(i - 16, i - 14) <> "EE") Then
                    .Range(i, i + 1).Selected
                Else
                    .Range(i + 1, i + 1).Selected
                End If
                .ForceEdit = blnForce
            End If
        End If
    End With

End Sub

Private Sub edt����_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, x As Single, y As Single)
    Dim Popup As CommandBar, Control As CommandBarControl
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_CUT, "����(&X)")
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "����(&C)")
        Set Control = .Add(xtpControlButton, ID_EDIT_PASTE, "ճ��(&V)    ")
        Popup.ShowPopup
    End With
End Sub

Private Sub Form_Activate()
    If blnActive = False Then
        If Me.txt���.Visible And Me.txt���.Enabled Then Me.txt���.SetFocus
        blnActive = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.tvw����.Visible Then
        Me.tvw����.Visible = False: Me.txt����.SetFocus: Exit Sub
    End If
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Me.edt����.PaperWidth = 16840
    Me.edt����.ResetWYSIWYG
    Set mfrmInsElement = New frmInsElement
    Set Elements = New cEPRElements
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Me.Width < Me.cmdOK.Left + Me.cmdOK.Width Then Me.Width = Me.cmdOK.Left + Me.cmdOK.Width
    If Me.Height < Me.pic����.Top + 2000 Then Me.Height = Me.pic����.Top + 2000
    With Me.pic����
        .Left = 0: .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Elements = Nothing
    Unload mfrmInsElement
    Set mfrmInsElement = Nothing
End Sub

Private Sub mfrmInsElement_pCancel()
    mfrmInsElement.Hide
    mfrmInsElement.Tag = ""
End Sub

Private Sub mfrmInsElement_pOK(Ele As cEPRElement)
    '��������Ҫ��
    Dim lngKey As Long
    If mfrmInsElement.Tag <> "" Then
        '�޸�ģʽ
        Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
        bInKeys = FindKey(Me.edt����, "E", mfrmInsElement.Tag, lSS, lSE, lES, lEE, bNeeded)
        If bInKeys Then
            Elements.Remove "K" & mfrmInsElement.Tag
            With Me.edt����
                .ForceEdit = True
                .Range(lSS, lEE) = ""
                .Range(lSS, lSS).Font.Protected = False
                .Range(lSS, lSS).Selected
                .ForceEdit = False
            End With
        End If
        lngKey = Elements.AddExistNode(Ele, True)
        Elements("K" & lngKey).InsertIntoEditor Me.edt����, , False
        bInKeys = FindKey(Me.edt����, "E", lngKey, lSS, lSE, lES, lEE, bNeeded)
        If bInKeys Then
            If Elements("K" & lngKey).������̬ = 0 Then
                Me.edt����.Range(lSE, lES).Selected
            Else
                Me.edt����.Range(lSE + 1, lSE + 1).Selected
            End If
        End If
    Else
        lngKey = Elements.AddExistNode(Ele)
        Elements("K" & lngKey).InsertIntoEditor Me.edt����, , True
    End If
    mfrmInsElement.Tag = ""
End Sub

Private Sub opt��Χ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub tvw����_DblClick()
    If Me.tvw����.SelectedItem Is Nothing Then Exit Sub
    Me.txt����.Tag = Mid(Me.tvw����.SelectedItem.Key, 2)
    Me.txt����.Text = Me.tvw����.SelectedItem.Text
    Me.txt����.SetFocus
    Call zlDefaultCode
End Sub

Private Sub tvw����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvw����.SelectedItem Is Nothing Then Exit Sub
        If Me.tvw����.SelectedItem.Children > 0 Then Exit Sub
        Call tvw����_DblClick
    Case vbKeySpace
        Call tvw����_DblClick
    Case vbKeyEscape
        Call tvw����_LostFocus
    End Select
End Sub

Private Sub tvw����_LostFocus()
    If Me.cmd���� Is ActiveControl Then Exit Sub
    Me.tvw����.Visible = False
End Sub

Private Sub txt���_Change()
    ValidControlText txt���
End Sub

Private Sub txt���_GotFocus()
    Me.txt���.SelStart = 0: Me.txt���.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt����_Change()
    ValidControlText txt����
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(Me.edt����.Text) = "" Then Me.edt����.Text = Trim(Me.txt����.Text)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    If InStr("%_'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

