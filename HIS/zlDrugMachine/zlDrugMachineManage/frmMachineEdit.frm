VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmMachineEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҩƷ�豸�༭"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8085
   Icon            =   "frmMachineEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   345
      Left            =   6840
      TabIndex        =   15
      Top             =   5430
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   345
      Left            =   5640
      TabIndex        =   14
      Top             =   5430
      Width           =   1095
   End
   Begin VB.CheckBox chkContine 
      Appearance      =   0  'Flat
      Caption         =   "���������豸�ӿ�(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Frame fraInfo 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton cmdLink 
         Caption         =   "��"
         Height          =   300
         Index           =   1
         Left            =   7440
         Picture         =   "frmMachineEdit.frx":038A
         TabIndex        =   16
         ToolTipText     =   "�������Ӵ�"
         Top             =   1200
         Width           =   330
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfInfo 
         Height          =   2415
         Left            =   360
         TabIndex        =   12
         Top             =   2160
         Width           =   7095
         _cx             =   12515
         _cy             =   4260
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.TextBox txtRemark 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   1680
         Width           =   6015
      End
      Begin VB.CommandButton cmdLink 
         Height          =   300
         Index           =   0
         Left            =   7125
         Picture         =   "frmMachineEdit.frx":13CC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "��������"
         Top             =   1200
         Width           =   330
      End
      Begin VB.TextBox txtLink 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   5670
      End
      Begin VB.ComboBox cboType 
         Appearance      =   0  'Flat
         Height          =   300
         ItemData        =   "frmMachineEdit.frx":240E
         Left            =   1440
         List            =   "frmMachineEdit.frx":2410
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   750
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   4
         Top             =   330
         Width           =   2175
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.ͬ�ⷿ��ͬ�ӿڵ�ҩƷ���Ͳ����ظ��� 2.ҩƷ���Ͳ��Ĭ��Ϊ����ҩƷ���ͣ�"
         Height          =   180
         Index           =   5
         Left            =   360
         TabIndex        =   17
         Top             =   4680
         Width           =   6570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע(&E)"
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   10
         Top             =   1710
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*��ַ(&I)"
         Height          =   180
         Index           =   3
         Left            =   360
         TabIndex        =   7
         Top             =   1230
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*�ӿ�����(&Y)"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   5
         Top             =   780
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*����(&M)"
         Height          =   180
         Index           =   1
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*���(&N)"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmMachineEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_BILL As String = "ҩƷ�ⷿ,,3,2000|�ⷿID,,0,0|ҩƷ����,,3,4500|���ͱ���,,0,0"

Private mblnShow As Boolean                     '��ʾ״̬��Load�¼���Ĺ��̴���
Private mblnReturn As Boolean                   '����ֵ�� Trueȷ�ϣ�Falseȡ��
Private mblnEdited As Boolean                   '�Ƿ��Ѿ��༭��True�ǣ�False��
Private mbytState As Byte                       '����״̬��0-�鿴��1-������2-�޸�
Private mlngID As Long                          'ҩƷ�豸�ӿڵ�ID
Private WithEvents mclsVSF As clsVSFlexGridEx
Attribute mclsVSF.VB_VarHelpID = -1

Public Function ShowMe(ByVal frmOwner As Form, ByVal bytState As Byte, Optional ByVal lngID As Long) As Boolean
'���ܣ�
'������
'  frmOwner�������������
'  bytState������״̬
'  lngID��ҩƷ�豸�ӿڵ�ID
'���أ�Trueȷ�ϣ�Falseȡ��

    If lngID = 0 And bytState <> Val("1-����") Then
        MsgBox "�봫��ӿ�ID��", vbInformation, GSTR_MSG
        Exit Function
    End If

    mbytState = bytState
    mlngID = lngID
    
    Me.Show vbModal, frmOwner
    ShowMe = mblnReturn

End Function

Private Sub cboType_Click()
    If Me.Visible = False Then Exit Sub
        
    Select Case Val(cboType.Text)
    Case Val("2-TOSHO"), Val("5-YUYAMA"), Val("6-��԰")
        lbl(3).Caption = "*���Ӵ�(&I)"
        txtLink.Locked = True
        txtLink.BackColor = &H8000000F
    Case Else
        lbl(3).Caption = "*��ַ(&I)"
        txtLink.Locked = False
        txtLink.BackColor = &H80000005
    End Select
    cmdLink(0).Visible = Not (Val(cboType.Text) = 2 Or Val(cboType.Text) = 5 Or Val(cboType.Text) = 6)
    cmdLink(1).Visible = (Val(cboType.Text) = 2 Or Val(cboType.Text) = 5 Or Val(cboType.Text) = 6)
    
    txtLink.Text = ""
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdLink_Click(Index As Integer)
    If Index = 0 Then
        'WebService
        Dim objSOAP As Object
        
        Call CreateSOAP(objSOAP)
        
        With objSOAP
            If Trim(txtLink.Text) = "" Then
                MsgBox "����д��" & cboType.Text & "����Ϣ��", vbInformation, GSTR_MSG
                If txtLink.Enabled Then txtLink.SetFocus
                Exit Sub
            End If
            
            On Error Resume Next
            .MSSoapInit txtLink.Text
            If Err.Number = 0 Then
                txtLink.Tag = "1"           '������ӳɹ�
                MsgBox "���ӳɹ���", vbInformation, GSTR_MSG
            Else
                txtLink.Tag = ""            '�������ʧ��
                If objSOAP Is Nothing Then
                    MsgBox "��SoapClient��δ��װ������ϵ������Ա��" & vbCrLf & _
                           "ע�⣺SoapClient��WinXP�°�װ2.0�汾��", _
                           vbInformation, GSTR_MSG
                Else
                    MsgBox "����ʧ�ܣ�", vbCritical, GSTR_MSG
                End If
            End If
            On Error GoTo 0
        End With
        Set objSOAP = Nothing
        
    Else
        'OLEDB���Ӵ�
        Dim msdLink As New MSDASC.DataLinks
        Dim cnTest As New ADODB.Connection
        
        If msdLink.PromptEdit(cnTest) Then
            On Error Resume Next
            Call cnTest.Open
            If Err.Number <> 0 Then
                txtLink.Text = ""
                txtLink.Tag = ""
                MsgBox "OLEDB���Ӵ�����ȷ�����飡", vbInformation, GSTR_MSG
            Else
                txtLink.Text = cnTest.ConnectionString
                txtLink.Tag = "1"
            End If
            cnTest.Close
            On Error GoTo 0
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    '���
    If Verify() = False Then Exit Sub
    
    '����
    If Save() = False Then Exit Sub
    
    If mbytState = enuEditState.���� And chkContine.Value Then
        '��������
        txtCode.Text = ""
        txtName.Text = ""
        txtRemark.Text = ""
        
        With vsfInfo
            .Redraw = False
            .Clear 1
            .Rows = 1
            .Redraw = True
        End With
        
        txtCode.SetFocus
    Else
        Unload Me
    End If
    
    mblnReturn = True
End Sub

Private Sub Form_Activate()
    If mblnShow Then
        Screen.MousePointer = vbHourglass
        
        Call InitControls
        If mbytState <> enuEditState.���� Then
            chkContine.Visible = False
            Call FillData
        End If
        
        mblnShow = False
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    mblnReturn = False
    mblnEdited = False
    
    Set mclsVSF = New clsVSFlexGridEx
    
    mblnShow = True         '���з����
End Sub

Private Sub InitControls()
    Dim arrTemp As Variant
    Dim i As Integer
    
    '�ؼ�λ��
    With cmdLink(1)
        .Left = cmdLink(0).Left
        .Top = cmdLink(0).Top
        .Width = cmdLink(0).Width
        .Height = cmdLink(0).Height
    End With

    '�ؼ�����ַ���
    mdlMain.SetTextMaxLen txtCode, "ҩƷ�豸�ӿ�.���"
    mdlMain.SetTextMaxLen txtName, "ҩƷ�豸�ӿ�.����"
    mdlMain.SetTextMaxLen txtLink, "ҩƷ�豸�ӿ�.������Ϣ"
    mdlMain.SetTextMaxLen txtRemark, "ҩƷ�豸�ӿ�.��ע"
    
    '��VSF
    With mclsVSF
        .Bunding = vsfInfo
        .Init
        .Head = MSTR_BILL
        .ColsReadonly = ""
        .Editable = EM_Modify
        .Repaint RT_Columns
    End With
    With vsfInfo
        .RowHeight(0) = 350
        .Rows = 2
        .ColComboList(.ColIndex("ҩƷ�ⷿ")) = "..."
        .ColComboList(.ColIndex("ҩƷ����")) = "..."
    End With
    
    '��ʼ�����Ϳؼ�����
    With cboType
        .Clear
        arrTemp = Split(GSTR_TYPE, "|")
        For i = LBound(arrTemp) To UBound(arrTemp)
            If arrTemp(i) <> "" Then
                .AddItem arrTemp(i)
            End If
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With
    Call cboType_Click
    
    '�������ÿؼ�Enabled����
    If mbytState = enuEditState.�鿴 Then
        For i = 0 To Me.Controls.Count - 1
            Select Case UCase(TypeName(Me.Controls(i)))
            Case "LABEL"
            Case Else
                Me.Controls(i).Enabled = False
            End Select
        Next
        cmdCancel.Enabled = True
    End If
    
End Sub

Private Sub FillData()
    Dim rsSQL As ADODB.Recordset, rsInfo As ADODB.Recordset
    Dim strInfo As String
    
    gstrSQL = "Select ID, ���, ����, ����, ������Ϣ, ��ע From ҩƷ�豸�ӿ� Where ID = [1] "
    
    On Error GoTo hErr
    Set rsSQL = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "ҩƷ�豸�ӿ�", mlngID)
    If rsSQL.EOF = False Then
        '���
        txtCode.Text = rsSQL!���
        '����
        txtName.Text = rsSQL!����
        '����
        cboType.ListIndex = rsSQL!���� - 1
        '���Ӵ���URL
        If IsNull(rsSQL!������Ϣ) Then
            txtLink.Text = ""
        Else
            txtLink.Text = mdlMain.Base64Decode(rsSQL!������Ϣ)
        End If
        '��ע
        txtRemark.Text = gobjComLib.zlCommFun.NVL(rsSQL!��ע)
        
        '�ⷿ�����
        gstrSQL = _
            "Select �ⷿID, '��' || �ⷿ���� || '��' || �ⷿ���� as ҩƷ�ⷿ " & vbNewLine & _
            "    , f_List2str(Cast(Collect(�������� Order By ���ͱ���) As t_Strlist), '��') ҩƷ����" & vbNewLine & _
            "    , f_List2str(Cast(Collect(���ͱ��� Order By ���ͱ���) As t_Strlist), '��') ���ͱ���" & vbNewLine & _
            "From (Select a.���� ���ͱ���, a.���� ��������, d.�ⷿid, b.���� �ⷿ����, b.���� �ⷿ����" & vbNewLine & _
            "      From ҩƷ���� A, ���ű� B, ҩƷ�豸�ӿ� C," & vbNewLine & _
            "         Xmltable('//root/bm' Passing c.��չ��Ϣ Columns �ⷿid Number(18) Path 'id', ���ͱ��� Varchar2(20) Path 'jxbm') D" & vbNewLine & _
            "      Where d.�ⷿid = b.Id(+) And d.���ͱ��� = a.����(+) And c.Id = [1] )" & vbNewLine & _
            "Group By �ⷿid, �ⷿ����, �ⷿ����" & vbNewLine & _
            "Union All " & vbNewLine & _
            "Select 0, '', '', '' From Dual "
        Set rsInfo = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "�ⷿ�����", mlngID)
        mclsVSF.Recordset = rsInfo
        mclsVSF.Repaint RT_Rows
        rsInfo.Close
        
        If vsfInfo.Rows <= 1 Then vsfInfo.Rows = 2
    End If
    rsSQL.Close
    
    Exit Sub
    
hErr:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVSF = Nothing
'    Set mfrmOwner = Nothing
End Sub

Private Sub txtCode_GotFocus()
    Call gobjComLib.zlControl.TXTSelAll(txtCode)
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    'ת��д
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then KeyAscii = KeyAscii - 32
    '����¼��
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub

Private Sub txtLink_Change()
    txtLink.Tag = ""
End Sub

Private Sub txtLink_GotFocus()
    Call gobjComLib.zlControl.TXTSelAll(txtLink)
End Sub

Private Sub txtLink_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtName_GotFocus()
    Call gobjComLib.zlControl.TXTSelAll(txtName)
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    '����¼��
    If InStr("`~!@#$%^&*()+={}|[]\:"";'<>?,./", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtRemark_GotFocus()
    Call gobjComLib.zlControl.TXTSelAll(txtRemark)
End Sub

Private Sub txtRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then gobjComLib.zlCommFun.PressKey vbKeyTab
End Sub

Private Function Verify() As Boolean
    Dim l As Long, lngID As Long, lngCount As Long
    Dim blnFind As Boolean
    
    '���
    If Trim(txtCode.Text) = "" Then
        MsgBox "����š�����δ��д��", vbInformation, GSTR_MSG
        txtCode.SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtCode.Text, vbFromUnicode)) > txtCode.MaxLength Then
        MsgBox mdlMain.FormatString("����š����ݳ��������[1]�ַ�����", txtCode.MaxLength), vbInformation, GSTR_MSG
        txtCode.SetFocus
        Exit Function
    End If
    If VerifyString(txtCode.Text, "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_") = False Then
        MsgBox "����š����ݴ��ڷǷ��ַ���", vbInformation, GSTR_MSG
        txtCode.SetFocus
        Exit Function
    End If
    
    '����
    If Trim(txtName.Text) = "" Then
        MsgBox "�����ơ�����δ��д��", vbInformation, GSTR_MSG
        txtName.SetFocus
        Exit Function
    End If
    If LenB(StrConv(txtName.Text, vbFromUnicode)) > txtName.MaxLength Then
        MsgBox mdlMain.FormatString("�����ơ����ݳ��������[1]�ַ�����", txtName.MaxLength), vbInformation, GSTR_MSG
        txtName.SetFocus
        Exit Function
    End If
    If VerifyString(txtName.Text, "`~!@#$%^&*()+={}|[]\:"";'<>?,./", False) = False Then
        MsgBox "�����ơ����ݴ��ڷǷ��ַ���", vbInformation, GSTR_MSG
        txtName.SetFocus
        Exit Function
    End If
    
    '��ַ
    If Trim(txtLink.Text) = "" Then
        If Val(cboType.Text) = Val("2-TOSHO") Then
            MsgBox mdlMain.FormatString("��[1]������δ���ã�", Split(lbl(3).Caption, "(")(0)), vbInformation, GSTR_MSG
            If cmdLink(1).Enabled And cmdLink(1).Visible Then cmdLink(1).SetFocus
        Else
            MsgBox mdlMain.FormatString("��[1]������δ��д��", Split(lbl(3).Caption, "(")(0)), vbInformation, GSTR_MSG
            txtLink.SetFocus
        End If
        Exit Function
    End If
    If LenB(StrConv(txtLink.Text, vbFromUnicode)) > txtLink.MaxLength Then
        MsgBox mdlMain.FormatString("��[1]�����ݳ��������[2]�ַ�����", Split(lbl(3).Caption, "(")(0), txtLink.MaxLength), _
                vbInformation, _
                GSTR_MSG
        txtLink.SetFocus
        Exit Function
    End If
    
    '��ע
    If Trim(txtRemark.Text) <> "" Then
        If LenB(StrConv(txtRemark.Text, vbFromUnicode)) > txtRemark.MaxLength Then
            MsgBox mdlMain.FormatString("����ע�����ݳ��������[1]�ַ�����", txtRemark.MaxLength), vbInformation, GSTR_MSG
            txtRemark.SetFocus
            Exit Function
        End If
        If VerifyString(txtRemark.Text, "`~!@#$%^&*()+={}|[]\:"";'<>?,./", False) = False Then
            MsgBox "����ע�����ݴ��ڷǷ��ַ���", vbInformation, GSTR_MSG
            txtRemark.SetFocus
            Exit Function
        End If
    End If
    
    '�ⷿ�����
    With vsfInfo
        For l = 1 To .Rows - 1
            lngID = Val(.TextMatrix(l, .ColIndex("�ⷿID")))
            If lngID > 0 Then
                blnFind = True
                Exit For
            End If
        Next
        If blnFind = False Then
            MsgBox "����дҩƷ�ⷿ��Ϣ��", vbInformation, GSTR_MSG
            Exit Function
        End If
    End With
        
    '���ⷿ��ҩƷ����
    With vsfInfo
        lngCount = 0
        blnFind = False
        For l = 1 To .Rows - 1
            lngID = Val(.TextMatrix(l, .ColIndex("�ⷿID")))
'            '��鵱ǰ�ӿڶ�ⷿ��ҩƷ���ͣ����������Ĭ�Ͽգ���ȫѡ��������
'            If lngID <> 0 Then
'                lngCount = lngCount + 1
'                If Trim(.TextMatrix(l, .ColIndex("���ͱ���"))) = "" Then
'                    blnFind = True
'                End If
'            End If
'            If lngCount > 0 And blnFind Then
'                MsgBox mdlMain.FormatString("��[1]���ĵ�ǰ�ӿ��Ѵ���ȫѡ��ҩƷ���ͣ�Ĭ�Ͽգ������飡", .TextMatrix(l, .ColIndex("ҩƷ�ⷿ")))
'                Exit Function
'            End If

            If Trim(.TextMatrix(l, .ColIndex("���ͱ���"))) = "" And lngID > 0 Then
                '�������ע��Ľӿ���������ҩƷ����Ϊ�յ�
                If CheckJiXing(lngID, Trim(txtCode.Text)) Then
                    MsgBox mdlMain.FormatString("��[1]���������ӿ���ȫѡҩƷ���ͣ����飡", .TextMatrix(l, .ColIndex("ҩƷ�ⷿ")))
                    Exit Function
                End If
            End If
        Next
        
    End With
    
    Verify = True
    
End Function

Private Function Save() As Boolean
    Dim strCode As String, strName As String, strType As String, strLink As String
    Dim strInfo As String, strRemark As String, strXml As String
    Dim arrJX As Variant
    Dim i As Long, j As Long
    Dim objXML As New clsXML
    
    strCode = "'" & Trim(txtCode.Text) & "'"
    strName = "'" & Trim(txtName.Text) & "'"
    strType = CStr(Val(cboType.Text))
    strLink = "'" & mdlMain.Base64Encode(Trim(txtLink.Text)) & "'"    '����
    strRemark = "'" & Trim(txtRemark.Text) & "'"
    
    '�ⷿ�����
    With vsfInfo
        objXML.ClearXmlText
        objXML.AppendNode "root", False
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("�ⷿID"))) > 0 Then
                If Trim(.TextMatrix(i, .ColIndex("���ͱ���"))) = "" Then
                    objXML.AppendNode "bm", False
                    objXML.AppendData "id", Trim(.TextMatrix(i, .ColIndex("�ⷿID")))
                    objXML.AppendData "jxbm", ""
                    objXML.AppendNode "bm", True
                Else
                    arrJX = Split(.TextMatrix(i, .ColIndex("���ͱ���")), "��")
                    For j = LBound(arrJX) To UBound(arrJX)
                        If arrJX(j) <> "" Then
                            objXML.AppendNode "bm", False
                            objXML.AppendData "id", Trim(.TextMatrix(i, .ColIndex("�ⷿID")))
                            objXML.AppendData "jxbm", arrJX(j)
                            objXML.AppendNode "bm", True
                        End If
                    Next
                End If
            End If
        Next
        objXML.AppendNode "root", True
    End With
    strInfo = "'" & Replace(Replace(objXML.XmlText, vbNewLine, ""), " ", "") & "'"
    
    Set objXML = Nothing
    
    Select Case mbytState
    Case enuEditState.����
        gstrSQL = mdlMain.FormatString("ZL_ҩƷ�豸�ӿ�_UPDATE([1], [2], [3], [4], [5], [6], [7])", _
                                        strCode, _
                                        strName, _
                                        strType, _
                                        strLink, _
                                        strInfo, _
                                        "Null", _
                                        strRemark)
    Case enuEditState.�޸�
        gstrSQL = mdlMain.FormatString("ZL_ҩƷ�豸�ӿ�_UPDATE([1], [2], [3], [4], [5], [6], [7])", _
                                        strCode, _
                                        strName, _
                                        strType, _
                                        strLink, _
                                        strInfo, _
                                        mlngID, _
                                        strRemark)
    End Select
    
    On Error GoTo hErr
    Call gobjComLib.zlDatabase.ExecuteProcedure(gstrSQL, "")
    
    Save = True
    Exit Function
    
hErr:
    Call gobjComLib.ErrCenter
End Function

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    If InStr("`~!@#$%^&*()+={}|[]\:"";'<>?,./", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsfInfo_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Row <= 0 Then Exit Sub
    
    Select Case Col
    Case vsfInfo.ColIndex("ҩƷ�ⷿ")
        Call Selector(1, Row)
        vsfInfo.TextMatrix(Row, vsfInfo.ColIndex("ҩƷ����")) = ""
        vsfInfo.TextMatrix(Row, vsfInfo.ColIndex("���ͱ���")) = ""
        
        If CheckDept() = False Then
            vsfInfo.TextMatrix(Row, vsfInfo.ColIndex("ҩƷ�ⷿ")) = ""
            vsfInfo.TextMatrix(Row, vsfInfo.ColIndex("�ⷿID")) = ""
        Else
            Call AppendSpaceLine
        End If
        
    Case vsfInfo.ColIndex("ҩƷ����")
        Call Selector(2, Row)
        
    End Select
End Sub

Private Sub AppendSpaceLine()
    Dim l As Long
    
    With vsfInfo
        l = .Rows - 1
        If Val(.TextMatrix(l, .ColIndex("�ⷿID"))) <> 0 Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        End If
    End With
End Sub

Private Function CheckDept() As Boolean
    Dim l As Long, lngID As Long
    Dim blnFind As Boolean
    Dim strCode As String
    
    
    With vsfInfo
        '���ⷿ�ظ�
        lngID = Val(.TextMatrix(.Row, .ColIndex("�ⷿID")))
        If lngID = 0 Then
            CheckDept = True
            Exit Function
        End If
        
        For l = 1 To .Rows - 1
            If lngID = Val(.TextMatrix(l, .ColIndex("�ⷿID"))) And l <> .Row Then
                blnFind = True
                Exit For
            End If
        Next
        If blnFind Then
            MsgBox "��ǰ��д�ġ�ҩƷ�ⷿ�����ظ������飡", vbInformation, GSTR_MSG
            Exit Function
        End If
        
    End With
    
    CheckDept = True
        
End Function

Private Function CheckJiXing(ByVal lngStoreID As Long, ByVal strInf As String) As Boolean
'���ܣ����ͬ�ⷿ�����ӿڵ�ҩƷ����
'������
'  lngStoreID���ⷿID
'  strInf���ӿڱ��
'���أ�True���ڣ�False������

    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo hErr
    
    gstrSQL = _
        "Select Count(1) Rec " & vbNewLine & _
        "From ҩƷ�豸�ӿ� A, Xmltable('//root/bm' Passing a.��չ��Ϣ Columns �ⷿid Number(18) Path 'id', ���ͱ��� Varchar2(20) Path 'jxbm') B " & vbNewLine & _
        "Where a.��� <> [1] And b.�ⷿid = [2] And b.���ͱ��� Is Null And Rownum < 2 "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "", strInf, lngStoreID)
    CheckJiXing = rsTmp!Rec > 0
    rsTmp.Close

    Exit Function
    
hErr:
    If gobjComLib.ErrCenter = 1 Then Resume
End Function

Private Sub Selector(ByVal intType As Integer, ByVal Row As Long)
'���ܣ�ѡ����
'������
'   intType��1-ҩƷ�ⷿ��2-ҩƷ����
'   Row��ѡ����ѡ�е�ֵҪд��ָ����

    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String, strCode As String
    Dim vRect As mdlDefine.RECT
    Dim sngTop As Single, sngLeft As Single
    Dim blnCancel As Boolean
    Dim lngDeptID As Long
    
    vRect = mdlMain.GetControlRect(vsfInfo.hwnd)
    sngTop = vRect.Top + vsfInfo.CellTop + vsfInfo.CellHeight
    sngLeft = vRect.Left + vsfInfo.CellLeft
    
    If intType = 1 Then
        '�ⷿ����ѡ��
        gstrSQL = _
            "Select Distinct a.Id, a.����, a.���� " & vbNewLine & _
            "From ���ű� A, ��������˵�� B " & vbNewLine & _
            "Where a.Id = b.����id And To_Char(Nvl(a.����ʱ��, To_Date('3000-1-1', 'yyyy-mm-dd')), 'yyyy') = '3000' " & vbNewLine & _
            "   And b.�������� In ('��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��') " & vbNewLine & _
            "Order By a.���� "
            
        Set rsTemp = gobjComLib.zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "ҩƷ�ⷿ", False, "", "ѡ��ҩƷ�ⷿ", False, False, True, _
                        sngLeft, sngTop, 0, blnCancel, False, False)
    Else
        '���ͣ���ѡ��
        lngDeptID = Val(vsfInfo.TextMatrix(Row, vsfInfo.ColIndex("�ⷿID")))
        
'        '��ǰ��ҩƷ����
'        strCode = BuildCode(vsfInfo.Row)
'        If strCode = "ALL" Then
'            MsgBox "ͬһ�ⷿ��������Ĭ��ȫѡҩƷ���ͣ��������", vbInformation, GSTR_MSG
'            Exit Sub
'        End If
        
        'ȡ��ҩƷ���ͣ����˵�ָ���ⷿ��ѡ��ҩƷ���ͣ����û�ѡ��
        gstrSQL = _
            "Select Rownum ID, ����, ���� " & vbNewLine & _
            "From ҩƷ���� " & vbNewLine & _
            "Where Not ���� In (Select b.���� " & vbNewLine & _
            "                   From ҩƷ�豸�ӿ� A," & vbNewLine & _
            "                      Xmltable('//root/bm' Passing a.��չ��Ϣ Columns �ⷿid Number(18) Path 'id', ���� Varchar2(20) Path 'jxbm') B " & vbNewLine & _
            "                   Where a.ID <> [2] And b.�ⷿid = [1] " & vbNewLine & _
            ") "

'        gstrSQL = _
'            "Select Rownum ID, ����, ���� " & vbNewLine & _
'            "From ҩƷ���� " & vbNewLine & _
'            "Order By ���� "

        Set rsTemp = gobjComLib.zlDatabase.ShowSQLMultiSelect(Me, gstrSQL, 0, "ҩƷ����", False, "����", "ѡ��ҩƷ����", False, False, True, _
                        sngLeft, sngTop, 0, blnCancel, False, False, lngDeptID, mlngID)
    End If
    
    If blnCancel = False Then
        If Not rsTemp Is Nothing Then
            With vsfInfo
                If intType = 1 Then
                    '�ⷿ
                    .TextMatrix(Row, .ColIndex("�ⷿID")) = rsTemp!ID
                    .TextMatrix(Row, .ColIndex("ҩƷ�ⷿ")) = "��" & rsTemp!���� & "��" & rsTemp!����
                Else
                    '����
                    strTemp = ""
                    strCode = ""
                    Do While rsTemp.EOF = False
                        strCode = strCode & "��" & rsTemp!����
                        strTemp = strTemp & "��" & rsTemp!����
                        rsTemp.MoveNext
                    Loop
                    If Left(strTemp, 1) = "��" Then strTemp = Mid(strTemp, 2)
                    If Left(strCode, 1) = "��" Then strCode = Mid(strCode, 2)
                    .TextMatrix(Row, .ColIndex("ҩƷ����")) = strTemp
                    .TextMatrix(Row, .ColIndex("���ͱ���")) = strCode
                End If
            End With
            rsTemp.Close
        Else
            If intType = 2 Then
                MsgBox "ͬһ�ⷿ�������ӿ�Ĭ��ȫѡҩƷ���ͣ��������", vbInformation, GSTR_MSG
            End If
        End If
    End If
End Sub

'Private Function BuildCode(ByVal lngCol As Long) As String
''���ܣ���ȡ��ǰ������ͬ�ⷿID��¼�ļ��ͣ�������ǰ��
''������
''  lngCol����ǰ��
'
'    Dim l As Long, lngDeptID As Long
'    Dim strCode As String, strTmp As String
'
'    With vsfInfo
'        lngDeptID = Val(.TextMatrix(lngCol, .ColIndex("�ⷿID")))
'
'        For l = 1 To .Rows - 1
'            'ͬ�ⷿID
'            If l <> lngCol And Val(.TextMatrix(l, .ColIndex("�ⷿID"))) = lngDeptID Then
'                strCode = Replace(Trim(.TextMatrix(l, .ColIndex("���ͱ���"))), "��", ",")
'                If strCode = "" Then
'                    '���м���
'                    BuildCode = "ALL"
'                    Exit Function
'                End If
'                strTmp = strTmp & "," & strCode
'            End If
'        Next
'        If Left(strTmp, 1) = "," Then strTmp = Mid(strTmp, 2)
'    End With
'
'    BuildCode = strTmp
'
'End Function

